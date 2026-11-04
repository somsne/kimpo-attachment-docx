<script setup lang="ts">
/**
 * 文档附件插件的显示区（kimpo-attachment-docx，填报态附件外壳内 / 历史记录页）：docx 的查看与编辑。
 *
 * 链路：宿主附件外壳按 manifest accepts 路由到本插件 → 挂载本组件（prop ctx: AttachmentOpenContext）。
 *   - 打开：ctx.readSource 读全文（宿主按 max_source_bytes = 50 MB 拒绝过大文件）→ 守门扫描（guard.ts）→
 *     查看态 DocxViewer（docx-preview 分页渲染，导航窗格「标题」按标题样式抽取）；
 *   - 守门命中（SmartArt / OLE 嵌入 / 图表 / 宏 / 文本框与形状 / 复杂域）→ 只读并出守门条幅（插件内 canEdit 降为假）；
 *   - 编辑：外壳切到编辑且 canEdit → 先取编辑锁（edit-flow.ts，与文本插件同一流程）→ 取到锁才挂 DocxEditor、
 *     按需加载 SuperDoc 引擎；引擎就绪后先做一次「导入自检」（原文件摘要 vs 编辑器模型），不一致转只读；
 *   - 保存：宿主标题栏「保存」经 ctx.onSaveRequest 回到插件 → 导出 docx → 重新解析做往返自检（编辑器模型 vs
 *     导出结果：段落数、表格数、正文文字去空白）→ 一致才 session.save（文件名与类型沿用原文件），否则拒存并 warn；
 *   - 失锁（EDIT_LOCK_LOST）：内容留在编辑器里，可「重新获取并保存」或（有下载权限时）下载副本。
 * 插件不屏蔽右键、选中与快捷键；水印由宿主叠加。AGPL-3.0：「关于」里给版本、许可与源码地址。
 */
import { computed, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue';
import type { AttachmentOpenContext } from '@kimposoft/plugin-sdk/attachment-open';
import Icon from './Icon.vue';
import DocxViewer from './DocxViewer.vue';
import DocxEditor from './DocxEditor.vue';
import { engineUrl, loadDocxEngine } from './engine-loader';
import { useAttachmentI18n } from './i18n';
import { useEditFlow } from './edit-flow';
import { scanDocx, type GuardResult } from './guard';
import { outlineOfDocx, type OutlineItem } from './outline';
import { compareSummaries, describeDiff, summarizeDocx, type DocxSummary } from './docx-summary';
import type { DocxEngine, DocxEngineModule } from './engine-types';

const props = defineProps<{
  ctx: AttachmentOpenContext;
  hostBridge?: unknown;
  /** 宿主给的随包资源基地址（#768）：引擎包 superdoc/engine.js 据此定位。 */
  assetBaseUrl?: string;
  /** 单测注入假引擎（生产不传）。 */
  loadEngine?: () => Promise<DocxEngineModule>;
}>();

/** 插件版本（构建时注入，单测缺省 dev）。 */
const PLUGIN_VERSION = typeof __PLUGIN_VERSION__ === 'string' ? __PLUGIN_VERSION__ : 'dev';
const SUPERDOC_VERSION = '1.47.2';
const SOURCE_REPO = 'https://github.com/somsne/kimpo-attachment-docx';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const { t, dispose: disposeI18n } = useAttachmentI18n(props.hostBridge, props.ctx.env?.locale);
/** 引擎包地址按宿主给的随包资源基地址定位（#768；缺键时 engineUrl 退回脚本地址）。 */
const loadEngineFromHost = () => loadDocxEngine(engineUrl(props.assetBaseUrl || undefined));

type Status = 'loading' | 'ready' | 'failed' | 'too_large';

const status = ref<Status>('loading');
/** 当前基线（打开时的原文件；保存成功后换成刚保存的内容）。查看态渲染它，编辑态从它开始。 */
const bytes = shallowRef<ArrayBuffer | null>(null);
const baseSummary = shallowRef<DocxSummary | null>(null);
const guard = shallowRef<GuardResult | null>(null);
const outline = shallowRef<OutlineItem[]>([]);
/** 引擎就绪后的导入自检没通过（转只读）。 */
const importCheckFailed = ref(false);
/** 最近一次保存的往返自检没通过。 */
const saveCheckFailed = ref(false);
const engine = shallowRef<DocxEngine | null>(null);
const aboutOpen = ref(false);
/** 编辑器重新挂载用的键（放弃修改后从基线重开）。 */
const editorKey = ref(0);
let pendingBytes: ArrayBuffer | null = null;
let disposed = false;

const canDownload = computed(() => props.ctx.permissions?.canDownload === true);
const locale = computed(() => props.ctx.env?.locale || 'zh-CN');
const timezone = computed(() => props.ctx.env?.timezone || '');
/** 当前用户名：契约暂无此字段，按可选字段读取，缺省「未署名」（批注与修订署名用）。 */
const userName = computed(() => {
  const loose = props.ctx as unknown as { user?: { displayName?: string }; env?: { userName?: string; userDisplayName?: string } };
  return loose.env?.userDisplayName || loose.env?.userName || loose.user?.displayName || t('docx.review.anonymous');
});

const guardText = computed(() => {
  const result = guard.value;
  if (!result?.readOnly) return '';
  const items = result.reasons.map(reason =>
    reason === 'complex_field'
      ? t('docx.guard.reason.complex_field', { fields: result.fieldTypes.join(' / ') })
      : t(`docx.guard.reason.${reason}`)
  );
  return t(canDownload.value ? 'docx.guard.banner' : 'docx.guard.banner_no_download', { items: items.join(t('docx.guard.separator')) });
});

const flow = useEditFlow({
  ctx: props.ctx,
  hostBridge: props.hostBridge,
  t,
  editable: () => status.value === 'ready' && !guard.value?.readOnly && !importCheckFailed.value,
  produce: produceChecked,
  onDiscard: () => {
    // 放弃修改：编辑器按基线重开（若仍在编辑态）
    editorKey.value++;
  },
  onSaved: () => {
    engine.value?.markSaved();
    if (pendingBytes) {
      bytes.value = pendingBytes;
      void summarizeDocx(pendingBytes).then(summary => {
        baseSummary.value = summary;
      });
      void outlineOfDocx(pendingBytes).then(items => {
        outline.value = items;
      });
    }
    pendingBytes = null;
  }
});

/** 编辑器显示中（取到锁之后；失锁后仍保留编辑器与内容）。 */
const editorShown = computed(() => flow.editing.value && !!bytes.value);
// 编辑器卸载（回到查看）时丢掉引擎引用
watch(editorShown, shown => {
  if (!shown) engine.value = null;
});

/** 编辑器内只读：未持锁（取锁中 / 失锁后可继续改，但这里只在导入自检失败时锁死）。 */
const editorReadOnly = computed(() => importCheckFailed.value);

/**
 * 生成保存内容：导出 → 重新解析 → 与编辑器模型比对（往返自检）。不一致拒存：上报 warn、出条幅、抛错
 * （edit-flow 原样 reject 给宿主），编辑器里的修改保留。
 */
async function produceChecked() {
  const current = engine.value;
  if (!current) throw Object.assign(new Error('docx editor not ready'), { code: 'EDITOR_NOT_READY' });
  const blob = await current.exportDocx();
  const exportedBytes = await blob.arrayBuffer();
  const expected = current.summarize();
  const actual = await summarizeDocx(exportedBytes);
  const diff = compareSummaries(expected, actual);
  if (!diff.ok) {
    saveCheckFailed.value = true;
    props.ctx.log('warn', '保存前往返自检未通过，已拒绝保存', {
      assetId: props.ctx.attachment.assetId,
      diff: describeDiff(diff),
      paragraphs: diff.paragraphs,
      tables: diff.tables,
      firstDiffIndex: diff.firstTextDiff?.index
    });
    throw Object.assign(new Error('docx round-trip check failed'), { code: 'ROUNDTRIP_CHECK_FAILED' });
  }
  saveCheckFailed.value = false;
  pendingBytes = exportedBytes;
  const type = props.ctx.attachment.contentType || DOCX_MIME;
  return { blob: new Blob([exportedBytes], { type }), contentType: type, name: props.ctx.attachment.name };
}

async function open(): Promise<void> {
  status.value = 'loading';
  props.ctx.setStatus({ state: 'loading' });
  try {
    const source = await props.ctx.readSource();
    if (disposed) return;
    const data = source.bytes;
    const [scanned, summary, items] = await Promise.all([
      scanDocx(data),
      summarizeDocx(data).catch(() => null),
      outlineOfDocx(data)
    ]);
    if (disposed) return;
    bytes.value = data;
    guard.value = scanned;
    baseSummary.value = summary;
    outline.value = items;
    importCheckFailed.value = false;
    status.value = 'ready';
    props.ctx.setStatus({ state: 'ready' });
    props.ctx.log('info', '文档附件已打开', { byteSize: data.byteLength, headings: items.length, guard: scanned.reasons });
    if (scanned.readOnly) props.ctx.log('warn', '文档含在线编辑不能无损保留的内容，按只读打开', { reasons: scanned.reasons, fields: scanned.fieldTypes });
    await flow.activate();
  } catch (error) {
    if (disposed) return;
    const code = (error as { code?: string } | null)?.code;
    status.value = code === 'SOURCE_TOO_LARGE' ? 'too_large' : 'failed';
    props.ctx.setStatus({ state: 'error', message: t(`docx.state.${status.value}`) });
    props.ctx.log(status.value === 'failed' ? 'error' : 'warn', '文档附件打开失败', { code, error: String(error) });
  }
}

/** 引擎就绪：导入自检（原文件摘要 vs 编辑器模型）。不一致说明编辑器没能完整装下原文件，转只读。 */
function onEditorReady(created: DocxEngine): void {
  engine.value = created;
  const base = baseSummary.value;
  if (!base) return;
  const diff = compareSummaries(base, created.summarize());
  if (!diff.ok) {
    importCheckFailed.value = true;
    created.setReadOnly(true);
    props.ctx.log('warn', '编辑器导入自检未通过，转为只读', { diff: describeDiff(diff), assetId: props.ctx.attachment.assetId });
  }
}

function onEditorChange(modified: boolean): void {
  if (!importCheckFailed.value) flow.markDirty(modified);
}

function onEditorError(): void {
  engine.value = null;
}

async function download(): Promise<void> {
  try {
    await props.ctx.download();
    props.ctx.log('info', '文档附件下载已交给宿主');
  } catch (error) {
    props.ctx.log('warn', '文档附件下载失败', { error: String(error) });
  }
}

/** 失锁 / 自检未过时，把编辑器里的内容存成本地副本（仅有下载权限时提供）。 */
async function downloadCopy(): Promise<void> {
  if (!engine.value) return;
  try {
    const blob = await engine.value.exportDocx();
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = props.ctx.attachment.name;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
    props.ctx.log('info', '下载了编辑中的本地副本', { byteSize: blob.size });
  } catch (error) {
    props.ctx.log('warn', '导出本地副本失败', { error: String(error) });
  }
}

function onViewerError(error: unknown): void {
  props.ctx.log('error', '文档渲染失败', { error: String(error) });
}

let offSourceChanged: (() => void) | undefined;

onMounted(() => {
  void open();
  // 附件在别处被换掉：没有未保存修改时重新读取；有修改时保留内容，保存会被宿主拒绝并提示
  offSourceChanged = props.ctx.onSourceChanged(() => {
    props.ctx.log('warn', '附件已在别处被替换', { dirty: flow.dirty.value });
    if (!flow.dirty.value) void open();
  });
});

onBeforeUnmount(() => {
  disposed = true;
  offSourceChanged?.();
  void flow.dispose();
  disposeI18n();
});

defineExpose({ flow, engine, status, guard });
</script>

<template>
  <div class="kad-root" data-testid="attachment-docx">
    <div v-if="guardText" class="kad-guard" role="note" data-testid="docx-guard">
      <Icon name="warning" />
      <span class="kad-grow">{{ guardText }}</span>
      <button v-if="canDownload" type="button" class="kad-banner-btn" data-testid="docx-guard-download" @click="download">{{ t('docx.guard.download') }}</button>
    </div>
    <div v-if="importCheckFailed" class="kad-guard" role="note" data-testid="docx-import-check">
      <Icon name="warning" />
      <span class="kad-grow">{{ t('docx.guard.selfcheck_open') }}</span>
      <button v-if="canDownload" type="button" class="kad-banner-btn" @click="download">{{ t('docx.guard.download') }}</button>
    </div>
    <div v-if="flow.lockHolder.value && !flow.locked.value" class="kad-guard" role="status" data-testid="docx-lock-held">
      <Icon name="lock" />
      <span class="kad-grow">{{ t('edit.lock_held_banner', { name: ctx.attachment.name, user: flow.lockHolder.value.displayName }) }}</span>
      <button v-if="flow.mode.value === 'edit'" type="button" class="kad-banner-btn" data-testid="docx-lock-retry" @click="flow.enterEdit()">{{ t('edit.retry') }}</button>
    </div>
    <div v-if="flow.lockLost.value" class="kad-alert" role="alert" data-testid="docx-lock-lost">
      <Icon name="alert" />
      <span class="kad-grow">{{ t('edit.lock_lost') }}</span>
      <button type="button" class="kad-banner-btn is-primary" :disabled="flow.saving.value || flow.acquiring.value" data-testid="docx-relock" @click="flow.relockAndSave()">{{ t('edit.relock_and_save') }}</button>
      <button v-if="canDownload" type="button" class="kad-banner-btn" data-testid="docx-download-copy" @click="downloadCopy">{{ t('edit.download_copy') }}</button>
    </div>
    <div v-if="saveCheckFailed" class="kad-alert" role="alert" data-testid="docx-save-check">
      <Icon name="alert" />
      <span class="kad-grow">{{ t('docx.save.check_failed') }}</span>
      <button v-if="canDownload" type="button" class="kad-banner-btn" @click="downloadCopy">{{ t('edit.download_copy') }}</button>
    </div>
    <div v-if="flow.errorKey.value" class="kad-alert" role="alert" data-testid="docx-error">
      <Icon name="alert" />
      <span class="kad-grow">{{ t(flow.errorKey.value) }}</span>
    </div>

    <div v-if="status !== 'ready'" class="kad-state" role="status" :data-state="status" data-testid="docx-state">
      <Icon :name="status === 'loading' ? 'spinner' : 'warning'" :size="28" :class="{ 'kad-spin': status === 'loading' }" />
      <span>{{ t(`docx.state.${status}`) }}</span>
      <button v-if="status === 'failed'" type="button" class="kad-text-btn" data-testid="docx-retry" @click="open">{{ t('docx.state.retry') }}</button>
    </div>
    <template v-else-if="bytes">
      <DocxEditor
        v-if="editorShown"
        :key="editorKey"
        :bytes="bytes"
        :file-name="ctx.attachment.name"
        :t="t"
        :locale="locale"
        :timezone="timezone"
        :user-name="userName"
        :track-changes="!!guard?.trackRevisions"
        :read-only="editorReadOnly"
        :can-download="canDownload"
        :log="ctx.log"
        :load-engine="loadEngine ?? loadEngineFromHost"
        @ready="onEditorReady"
        @change="onEditorChange"
        @error="onEditorError"
        @download="download"
        @about="aboutOpen = true"
      />
      <DocxViewer
        v-else
        :bytes="bytes"
        :outline="outline"
        :can-download="canDownload"
        :t="t"
        @download="download"
        @about="aboutOpen = true"
        @error="onViewerError"
      />
    </template>

    <div v-if="aboutOpen" class="kad-dialog-mask" @click.self="aboutOpen = false">
      <div class="kad-dialog" role="dialog" aria-modal="true" :aria-label="t('docx.about.title')" data-testid="docx-about-dialog">
        <div class="kad-dialog-title">{{ t('docx.about.title') }}</div>
        <p>{{ t('docx.about.version', { v: PLUGIN_VERSION }) }}</p>
        <p>{{ t('docx.about.license', { sd: SUPERDOC_VERSION }) }}</p>
        <p>{{ t('docx.about.source') }}：<a :href="SOURCE_REPO" target="_blank" rel="noopener noreferrer" data-testid="docx-about-source">{{ SOURCE_REPO }}</a></p>
        <p>{{ t('docx.about.offer') }}</p>
        <div class="kad-dialog-actions">
          <button type="button" class="kad-text-btn" @click="aboutOpen = false">{{ t('docx.about.close') }}</button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.kad-root {
  position: relative;
  display: flex;
  flex-direction: column;
  width: 100%;
  height: 100%;
  min-height: 0;
  color: var(--kad-text);
  background: var(--kad-bg);
}

.kad-dialog a {
  color: var(--kad-primary);
  word-break: break-all;
}
</style>
