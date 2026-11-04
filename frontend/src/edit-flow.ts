/**
 * 附件插件的编辑流（文档附件插件 kimpo-attachment-docx，填报态附件外壳内；照搬自文本附件插件 kimpo-attachment-text 的 edit-flow.ts，改动须两边同步）。
 *
 * 契约见 SDK attachment-open.ts：宿主管权限、编辑锁、心跳与保存后经表单数据权威层（代号 艾莉亚(Aria)）写回字段，
 * 插件只管内容。本模块把插件侧的流程收在一处：
 *   - 外壳切到编辑（ctx.mode / onModeChange 为 'edit'）且 permissions.canEdit 与插件自判都允许时，先 ctx.beginEdit() 取锁，
 *     取到锁才放开编辑；被占用（EDIT_LOCKED）时用宿主确认框提示占用者（重试 / 仅查看），没有确认框能力时只出条幅；
 *   - 有改动 → ctx.setDirty(true)（点亮宿主标题栏的「保存」）；
 *   - 取到锁（进入编辑态）后登记一次 ctx.onSaveRequest：宿主标题栏「保存」「保存并关闭」、结束编辑前的确认都经它回到插件，
 *     回调里 session.save(blob, {contentType, name}) 写出——resolve true = 已保存、false = 没有可保存的改动，
 *     reject = 保存失败（宿主按 error.code 提示，插件保留内容）；成功 → setDirty(false) 并上报 info 日志；
 *   - 保存时失锁（EDIT_LOCK_LOST）→ 保留内容、出条幅，可「重新获取并保存」；宿主再请保存时先静默重新取锁；
 *   - 外壳切回查看（宿主已先经 onSaveRequest 让用户决定保存或放弃）→ 结束会话，仍有未保存修改即按放弃结束并恢复原内容。
 * 插件不屏蔽右键、选中与快捷键。
 */
import { ref, type Ref } from 'vue';
import type {
  AttachmentEditLockHolder,
  AttachmentEditSession,
  AttachmentOpenContext,
  AttachmentOpenError,
  AttachmentOpenMode
} from '@kimposoft/plugin-sdk/attachment-open';
import { createHostUi } from '@kimposoft/plugin-sdk/ui';

export type Translate = (key: string, params?: Record<string, unknown>) => string;

export interface SavePayload {
  blob: Blob;
  contentType?: string;
  name?: string;
}

export interface EditFlowOptions {
  ctx: AttachmentOpenContext;
  hostBridge: unknown;
  t: Translate;
  /** 插件自判能否编辑（如内容被截断、格式不支持时为 false）。 */
  editable: () => boolean;
  /** 生成保存内容。 */
  produce: () => Promise<SavePayload> | SavePayload;
  /** 放弃未保存修改：把内容恢复为最近一次保存的基线。 */
  onDiscard: () => void;
  /** 保存成功：把当前内容记为新基线。 */
  onSaved: () => void;
}

/** 编辑流对外状态（模板与测试读取）。 */
export interface EditFlow {
  mode: Ref<AttachmentOpenMode>;
  /** 持有编辑锁（可以编辑）。失锁后内容仍可改，但 locked 为 false，保存前需重新取锁。 */
  locked: Ref<boolean>;
  /** 进入过编辑且尚未回到查看（编辑工具可用）。 */
  editing: Ref<boolean>;
  acquiring: Ref<boolean>;
  saving: Ref<boolean>;
  dirty: Ref<boolean>;
  /** 编辑锁被占用时的占用者（出条幅）。 */
  lockHolder: Ref<AttachmentEditLockHolder | null>;
  /** 保存时失锁。 */
  lockLost: Ref<boolean>;
  /** 插件自己发起的操作（取锁、条幅上的「重新获取并保存」）最近一次失败的提示文案 key（null = 无）。 */
  errorKey: Ref<string | null>;
  /** 当前能否进入编辑（权限 + 插件自判）。 */
  canEdit(): boolean;
  /** 内容就绪后调用：外壳已处于编辑态时立即取锁。 */
  activate(): Promise<void>;
  markDirty(dirty: boolean): void;
  enterEdit(): Promise<boolean>;
  /** 宿主经 onSaveRequest 请插件保存时的处理（导出供测试）。 */
  handleSaveRequest(): Promise<boolean>;
  relockAndSave(): Promise<boolean>;
  dispose(): Promise<void>;
}

function errorOf(error: unknown): Partial<AttachmentOpenError> {
  return (error ?? {}) as Partial<AttachmentOpenError>;
}

/** UTC ISO 时间按用户时区显示为 YYYY-MM-DD HH:mm；时区名无效时退回浏览器本地时间。 */
export function formatLockTime(iso: string, timeZone: string): string {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return iso;
  const pad = (n: number) => String(n).padStart(2, '0');
  let stamp = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
  try {
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
    }).formatToParts(date);
    const part = (type: Intl.DateTimeFormatPartTypes) => parts.find(p => p.type === type)?.value ?? '';
    stamp = `${part('year')}-${part('month')}-${part('day')} ${part('hour')}:${part('minute')}`;
  } catch {
    // 时区名无效：保留浏览器本地时间
  }
  return stamp;
}

const SAVE_ERROR_KEYS: Record<string, string> = {
  SAVE_TOO_LARGE: 'edit.error.save_too_large',
  SAVE_TYPE_NOT_ALLOWED: 'edit.error.save_type_not_allowed',
  EDIT_FORBIDDEN: 'edit.error.forbidden',
  SOURCE_CHANGED: 'edit.error.source_changed'
};

export function useEditFlow(options: EditFlowOptions): EditFlow {
  const { ctx, t } = options;
  const ui = createHostUi(options.hostBridge);
  const mode = ref<AttachmentOpenMode>(ctx.mode === 'edit' ? 'edit' : 'view');
  const locked = ref(false);
  const editing = ref(false);
  const acquiring = ref(false);
  const saving = ref(false);
  const dirty = ref(false);
  const lockHolder = ref<AttachmentEditLockHolder | null>(null);
  const lockLost = ref(false);
  const errorKey = ref<string | null>(null);
  let session: AttachmentEditSession | null = null;
  /** 最近一次取锁失败的错误（宿主请保存而重新取锁失败时原样 reject 给宿主）。 */
  let lastBeginError: unknown = null;
  /** 正在进行的写出（宿主连按保存时复用同一次）。 */
  let inflight: Promise<void> | null = null;
  /** onSaveRequest 的取消订阅；进入编辑态后登记一次，回到查看或卸载时取消。 */
  let offSave: (() => void) | null = null;
  let active = false;
  let disposed = false;

  const canEdit = () => ctx.permissions?.canEdit === true && options.editable();

  function markDirty(next: boolean): void {
    if (dirty.value === next) return;
    dirty.value = next;
    ctx.setDirty(next);
  }

  /** 被占用：有宿主确认框时问「重试 / 仅查看」，选重试则再取一次锁。 */
  async function promptLockHeld(holder: AttachmentEditLockHolder | null): Promise<boolean> {
    lockHolder.value = holder;
    if (!ui.has('confirm')) return false;
    try {
      const result = await ui.openConfirm({
        type: 'warning',
        title: t('edit.lock_held_title'),
        content: t('edit.lock_held_body', {
          name: ctx.attachment.name,
          user: holder?.displayName ?? '',
          time: holder?.since ? formatLockTime(holder.since, ctx.env?.timezone ?? '') : ''
        }),
        positiveText: t('edit.retry'),
        negativeText: t('edit.view_only')
      });
      return result.confirmed;
    } catch (error) {
      ctx.log('warn', '编辑锁占用确认框打开失败', { error: String(error) });
      return false;
    }
  }

  async function enterEdit(): Promise<boolean> {
    if (disposed || locked.value || acquiring.value || !canEdit()) return locked.value;
    acquiring.value = true;
    errorKey.value = null;
    try {
      for (;;) {
        try {
          session = await ctx.beginEdit();
          if (disposed || mode.value !== 'edit') {
            // 取锁期间外壳已切回查看或组件已卸载：立即还锁
            await session.end({ discard: true }).catch(() => undefined);
            session = null;
            return false;
          }
          locked.value = true;
          editing.value = true;
          lockHolder.value = null;
          lockLost.value = false;
          lastBeginError = null;
          offSave ??= ctx.onSaveRequest(handleSaveRequest);
          ctx.log('info', '进入编辑', { assetId: ctx.attachment.assetId, sessionId: session.id });
          return true;
        } catch (error) {
          lastBeginError = error;
          const err = errorOf(error);
          if (err.code === 'EDIT_LOCKED') {
            ctx.log('info', '附件正被他人编辑', { lockedBy: err.data?.lockedBy?.displayName });
            if (await promptLockHeld(err.data?.lockedBy ?? null)) continue;
            return false;
          }
          errorKey.value = SAVE_ERROR_KEYS[err.code ?? ''] ?? 'edit.error.begin_failed';
          ctx.log('warn', '进入编辑失败', { code: err.code, error: String(error) });
          return false;
        }
      }
    } finally {
      acquiring.value = false;
    }
  }

  /** 写出一次：失败抛错；失锁时清掉会话、出条幅（内容保留）。 */
  function writeOut(): Promise<void> {
    inflight ??= (async () => {
      const current = session;
      if (!current) throw Object.assign(new Error('attachment edit lock lost'), { code: 'EDIT_LOCK_LOST' });
      saving.value = true;
      try {
        const payload = await options.produce();
        const result = await current.save(payload.blob, { contentType: payload.contentType, name: payload.name });
        options.onSaved();
        markDirty(false);
        lockLost.value = false;
        errorKey.value = null;
        ctx.log('info', '附件已保存', { assetId: result.assetId, byteSize: result.byteSize });
      } catch (error) {
        const err = errorOf(error);
        if (err.code === 'EDIT_LOCK_LOST') {
          session = null;
          locked.value = false;
          lockLost.value = true;
          ctx.log('warn', '保存时编辑锁已失效，内容保留在页面上', { assetId: ctx.attachment.assetId });
        } else {
          ctx.log('error', '附件保存失败', { code: err.code, error: String(error) });
        }
        throw error;
      } finally {
        saving.value = false;
      }
    })().finally(() => {
      inflight = null;
    });
    return inflight;
  }

  /**
   * 宿主请保存（onSaveRequest 回调）：没有改动 → false；失锁后先静默重新取锁，取不到则把取锁错误 reject 给宿主；
   * 写出失败原样 reject（宿主按 error.code 提示），插件只保留内容与失锁条幅，不另出错误条幅。
   */
  async function handleSaveRequest(): Promise<boolean> {
    if (inflight) await inflight;
    if (!dirty.value) return false;
    if (!session && !(await enterEdit())) {
      throw lastBeginError ?? Object.assign(new Error('attachment edit lock lost'), { code: 'EDIT_LOCK_LOST' });
    }
    await writeOut();
    return true;
  }

  /** 条幅上的「重新获取并保存」：插件自己发起，失败在插件内出条幅。 */
  async function relockAndSave(): Promise<boolean> {
    if (!(await enterEdit())) return false;
    if (!dirty.value) return true;
    try {
      await writeOut();
      return true;
    } catch (error) {
      const code = errorOf(error).code ?? '';
      if (code !== 'EDIT_LOCK_LOST') errorKey.value = SAVE_ERROR_KEYS[code] ?? 'edit.error.save_failed';
      return false;
    }
  }

  /** 回到查看：结束会话；有未保存修改时外壳已确认放弃，恢复基线。 */
  async function leaveEdit(): Promise<void> {
    const wasDirty = dirty.value;
    const current = session;
    session = null;
    offSave?.();
    offSave = null;
    locked.value = false;
    editing.value = false;
    lockLost.value = false;
    lockHolder.value = null;
    errorKey.value = null;
    if (wasDirty) {
      options.onDiscard();
      markDirty(false);
    }
    if (current) {
      try {
        await current.end({ discard: wasDirty });
        ctx.log('info', '结束编辑', { discarded: wasDirty });
      } catch (error) {
        ctx.log('warn', '结束编辑失败', { error: String(error) });
      }
    }
  }

  const offMode = ctx.onModeChange(next => {
    mode.value = next;
    if (next === 'edit') {
      if (active) void enterEdit();
    } else {
      void leaveEdit();
    }
  });

  async function activate(): Promise<void> {
    active = true;
    if (mode.value === 'edit') await enterEdit();
  }

  async function dispose(): Promise<void> {
    disposed = true;
    offMode?.();
    offSave?.();
    offSave = null;
    const current = session;
    session = null;
    if (current) await current.end({ discard: dirty.value }).catch(() => undefined);
  }

  return {
    mode,
    locked,
    editing,
    acquiring,
    saving,
    dirty,
    lockHolder,
    lockLost,
    errorKey,
    canEdit,
    activate,
    markDirty,
    enterEdit,
    handleSaveRequest,
    relockAndSave,
    dispose
  };
}
