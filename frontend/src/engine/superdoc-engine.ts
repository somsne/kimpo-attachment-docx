/**
 * SuperDoc 编辑引擎封装（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 编辑态）。
 *
 * 实现 engine-types.ts 的 DocxEngine：界面（DocxEditor.vue）只经接口驱动，不直接碰 SuperDoc。
 *   - SuperDoc 锁 1.47.2（AGPL-3.0）：2.x 起依赖专有闭源的 @superdoc/docx-engine（禁止再分发），不能随插件发布；
 *   - 不用 SuperDoc 自带工具栏与批注栏皮肤（不传 toolbar 选择器、批注用 headless 控制器自绘），功能区按设计稿自绘；
 *   - 遥测显式关闭（SuperDoc 默认 enabled，会向 ingest.superdoc.dev 发数据；插件声明无网络权限）；
 *   - 导出直接调主编辑器 exportDocx：superdoc.export() 失败时会悄悄回落原文件，保存自检就看不出丢了修改。
 * 本文件随 engine.js 单独构建（vite.engine.config.ts），进入编辑态才加载。
 */
import { SuperDoc } from 'superdoc';
import { createSuperDocUI } from 'superdoc/ui';
import type { DocxSummary } from '../docx-summary';
import { summarizeEditorDoc, type PmNodeLike } from '../docx-summary';
import type {
  CommandState,
  DocxCommandId,
  DocxEngine,
  DocxEngineOptions,
  EngineChange,
  EngineComment,
  EngineOutlineItem,
  EngineSnapshot,
  FindResult,
  PageSetup,
  StyleOption
} from '../engine-types';
import { countWords } from '../word-count';

/** SuperDoc 内置、可直接转交 ui.toolbar.execute 的命令。 */
const BUILT_IN: ReadonlySet<string> = new Set([
  'bold', 'italic', 'underline', 'strikethrough', 'font-size', 'font-family', 'text-color', 'highlight-color', 'link',
  'text-align', 'line-height', 'bullet-list', 'numbered-list', 'indent-increase', 'indent-decrease', 'undo', 'redo',
  'ruler', 'formatting-marks', 'zoom', 'zoom-fit-width', 'clear-formatting', 'copy-format',
  'track-changes-accept-selection', 'track-changes-reject-selection', 'image', 'table-of-contents-insert', 'table-insert',
  'table-add-row-before', 'table-add-row-after', 'table-delete-row', 'table-add-column-before', 'table-add-column-after',
  'table-delete-column', 'table-delete', 'table-merge-cells', 'table-split-cell'
]);

/** 纸张尺寸（英寸，纵向宽×高）。 */
const PAPER: Record<string, [number, number]> = {
  A4: [8.2681, 11.6929],
  A3: [11.6929, 16.5354],
  B5: [6.9291, 9.8425],
  Letter: [8.5, 11],
  Legal: [8.5, 14]
};

/** 字号阶梯（pt），「增大 / 减小字号」按它跳。 */
const SIZE_STEPS = [5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 15, 16, 18, 20, 22, 24, 26, 28, 36, 42, 48, 72];

/* eslint-disable @typescript-eslint/no-explicit-any -- SuperDoc 1.x 的编辑器实例类型面很宽，这里按运行时形状取用 */
type AnyEditor = any;
type AnyUi = any;

function toBlob(out: unknown): Blob {
  if (out instanceof Blob) return out;
  if (out instanceof ArrayBuffer || ArrayBuffer.isView(out)) {
    return new Blob([out as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }
  throw Object.assign(new Error('docx export returned nothing'), { code: 'EXPORT_EMPTY' });
}

function isoOf(value: unknown): string | null {
  if (typeof value === 'number' && Number.isFinite(value)) return new Date(value).toISOString();
  if (typeof value === 'string' && value) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? null : date.toISOString();
  }
  return null;
}

/** 样式名是否 Word 预置（用于界面翻译内置名、标出自定义样式）。 */
export function isBuiltinStyleName(name: string): boolean {
  return /^(normal|title|subtitle|quote|intense quote|heading [1-9]|list paragraph|no spacing|caption|toc [1-9]|footnote text|endnote text|header|footer|strong|emphasis)$/i.test(name.trim());
}

function headingLevelOf(style: any, all: Map<string, any>): number {
  let current = style;
  for (let depth = 0; current && depth < 10; depth++) {
    const attrs = current.definition?.attrs ?? {};
    if (Number.isInteger(attrs.outlineLevel) && attrs.outlineLevel >= 0 && attrs.outlineLevel < 9) return attrs.outlineLevel + 1;
    const byName = /^heading\s*([1-9])$/i.exec(String(attrs.name ?? '').trim());
    if (byName) return Number(byName[1]);
    if (/^title$/i.test(String(attrs.name ?? '').trim())) return 1;
    current = attrs.basedOn ? all.get(attrs.basedOn) : null;
  }
  return 0;
}

function ptOf(value: unknown): number | undefined {
  const n = parseFloat(String(value ?? ''));
  return Number.isFinite(n) && n > 0 ? n : undefined;
}

export function createDocxEngine(options: DocxEngineOptions): DocxEngine {
  const { log } = options;
  const listeners = new Set<(snapshot: EngineSnapshot) => void>();
  const contentListeners = new Set<(modified: boolean) => void>();
  /** 基线：载入完成（或上次保存）时的正文与批注签名，判断有无未保存修改用。 */
  let baselineDoc: any = null;
  let baselineComments = '';
  let modified = false;
  /** 用户是否已经动过编辑区（键盘、指针、粘贴、拖放）；之前的任何变化都算载入期规范化。 */
  let userTouched = false;
  let editor: AnyEditor = null;
  let ui: AnyUi = null;
  let destroyed = false;
  let readyFlag = false;
  let trackChanges = options.trackChanges;
  let readOnly = false;
  let pages = 1;
  let currentPage = 1;
  let words = 0;
  let wordTimer: ReturnType<typeof setTimeout> | null = null;
  let notifyQueued = false;
  let search: { query: string; caseSensitive: boolean; matches: any[]; index: number } | null = null;
  let scrollEl: HTMLElement | null = null;
  const cleanups: Array<() => void> = [];

  const host = document.createElement('div');
  host.className = 'kad-superdoc-host';
  options.element.appendChild(host);

  let resolveReady!: () => void;
  let rejectReady!: (error: unknown) => void;
  const ready = new Promise<void>((resolve, reject) => {
    resolveReady = resolve;
    rejectReady = reject;
  });

  const file = new File([options.file], options.fileName || 'document.docx', {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });

  const superdoc: any = new SuperDoc({
    selector: host,
    document: file,
    documentMode: trackChanges ? 'suggesting' : 'editing',
    role: 'editor',
    user: { name: options.user.name || 'Kimpo', email: options.user.email ?? '' },
    telemetry: { enabled: false },
    rulers: true,
    contained: true,
    disablePiniaDevtools: true,
    // SuperDoc 自带右键菜单是英文界面，关掉，由插件自绘四语右键菜单（DocxEditor.vue）
    disableContextMenu: true,
    modules: {
      comments: { displayMode: 'inline' },
      toolbar: { hideButtons: true }
    },
    onEditorCreate: ({ editor: created }: { editor: AnyEditor }) => {
      // 主体编辑器只认第一个（页眉页脚编辑器不覆盖）
      if (!editor && !created?.options?.isHeaderOrFooter) editor = created;
    },
    onReady: () => {
      if (destroyed) return;
      editor ??= superdoc.activeEditor;
      if (!editor) {
        rejectReady(new Error('docx editor not created'));
        return;
      }
      ui = createSuperDocUI({ superdoc });
      for (const slice of ['toolbar', 'selection', 'comments', 'trackChanges', 'zoom', 'fonts']) {
        const off = ui[slice]?.observe?.(() => queueNotify());
        if (typeof off === 'function') cleanups.push(off);
      }
      editor.on?.('update', onDocUpdate);
      editor.on?.('selectionUpdate', queueNotify);
      cleanups.push(() => {
        editor?.off?.('update', onDocUpdate);
        editor?.off?.('selectionUpdate', queueNotify);
      });
      attachScroll();
      recountWords();
      adoptHiddenHost();
      readyFlag = true;
      rebaseline();
      log('info', '文档编辑引擎就绪', { trackChanges, pages });
      resolveReady();
      queueNotify();
    },
    onPaginationUpdate: ({ totalPages }: { totalPages: number }) => {
      pages = Math.max(1, totalPages || 1);
      attachScroll();
      queueNotify();
    },
    onCommentsUpdate: () => {
      // 批注的增删改与回复不一定改正文（回复只进 comments.xml），按指纹比对
      if (readyFlag) queueMicrotask(emitContent);
      queueNotify();
    },
    onContentError: ({ error }: { error: unknown }) => {
      log('error', '文档导入失败', { error: String(error) });
      rejectReady(error);
    },
    onException: (payload: unknown) => {
      log('warn', 'SuperDoc 异常', { error: String((payload as { error?: unknown })?.error ?? payload) });
    }
  });

  /**
   * #794：SuperDoc 的演示式编辑器（PresentationEditor）把真正可编辑的 ProseMirror 放在 document.body 下的隐藏宿主
   * （.presentation-editor__hidden-host-wrapper，fixed 定位到屏幕外），正文显示层只负责画。它在附件外壳的 DOM 之外，
   * 外壳的焦点陷阱会把落到它上面的焦点抢回外壳（功能区第一个可聚焦元素），正文输入与拖选随之失效。
   * 把这个宿主搬进插件自己的挂载容器：它仍是 fixed + 屏幕外，SuperDoc 按 DOM 引用使用它，搬家不影响渲染，
   * 但从此它在外壳之内、焦点陷阱认它为内部。就绪与每次状态通知时各核一次，SuperDoc 重建宿主也能跟上。
   */
  function adoptHiddenHost(): void {
    if (destroyed) return;
    const dom = editor?.view?.dom as HTMLElement | undefined;
    const wrapper = dom?.closest?.('.presentation-editor__hidden-host-wrapper') as HTMLElement | null | undefined;
    if (!wrapper || wrapper.parentElement === host) return;
    host.appendChild(wrapper);
  }

  function commentsSignature(): string {
    const items: any[] = safe(() => ui?.comments?.getSnapshot?.()?.items ?? [], []);
    return JSON.stringify(items.map(c => [c.id, c.text, c.status, c.parentCommentId ?? null]));
  }

  function rebaseline(): void {
    baselineDoc = editor?.state?.doc ?? null;
    baselineComments = commentsSignature();
    if (modified) {
      modified = false;
      contentListeners.forEach(fn => fn(false));
    }
  }

  /** 内容可能变了：用户动过之前一律重设基线；之后与基线比对，「有无修改」翻转时通知。 */
  function emitContent(): void {
    if (!readyFlag || !editor) return;
    if (!userTouched) {
      rebaseline();
      return;
    }
    const next = !(baselineDoc && editor.state.doc.eq(baselineDoc)) || commentsSignature() !== baselineComments;
    if (next !== modified) {
      modified = next;
      contentListeners.forEach(fn => fn(next));
    }
  }

  const markTouched = () => {
    userTouched = true;
  };
  for (const type of ['keydown', 'pointerdown', 'paste', 'drop', 'cut']) host.addEventListener(type, markTouched, true);
  cleanups.push(() => {
    for (const type of ['keydown', 'pointerdown', 'paste', 'drop', 'cut']) host.removeEventListener(type, markTouched, true);
  });

  function onDocUpdate(): void {
    if (!readyFlag) return;
    search = search ? { ...search, matches: [], index: -1 } : null;
    emitContent();
    if (wordTimer) clearTimeout(wordTimer);
    wordTimer = setTimeout(recountWords, 300);
    queueNotify();
  }

  function recountWords(): void {
    wordTimer = null;
    if (!editor) return;
    try {
      words = countWords(summarize().paragraphTexts.join('\n'));
    } catch {
      // 统计失败不影响编辑
    }
    queueNotify();
  }

  function pageElements(): HTMLElement[] {
    return Array.from(host.querySelectorAll<HTMLElement>('.superdoc-page, [data-page-index]'));
  }

  function updateCurrentPage(): void {
    const list = pageElements();
    if (!list.length || !scrollEl) return;
    const middle = scrollEl.getBoundingClientRect().top + scrollEl.clientHeight / 3;
    let page = 1;
    list.forEach((el, i) => {
      if (el.getBoundingClientRect().top <= middle) page = i + 1;
    });
    if (page !== currentPage) {
      currentPage = page;
      queueNotify();
    }
  }

  function attachScroll(): void {
    const el: HTMLElement | null = ui?.viewport?.getScrollContainer?.() ?? host.querySelector('.superdoc') ?? host;
    if (!el || el === scrollEl) return;
    scrollEl?.removeEventListener('scroll', updateCurrentPage);
    scrollEl = el;
    el.addEventListener('scroll', updateCurrentPage, { passive: true });
  }
  cleanups.push(() => scrollEl?.removeEventListener('scroll', updateCurrentPage));

  function queueNotify(): void {
    adoptHiddenHost();
    if (notifyQueued || destroyed) return;
    notifyQueued = true;
    queueMicrotask(() => {
      notifyQueued = false;
      if (destroyed) return;
      const snapshot = getSnapshot();
      listeners.forEach(fn => fn(snapshot));
    });
  }

  function marksAtSelection(): any[] {
    try {
      const { selection } = editor.state;
      return selection.empty ? selection.$from.marks() : [];
    } catch {
      return [];
    }
  }

  function vertAlignActive(kind: 'superscript' | 'subscript'): boolean {
    return marksAtSelection().some((mark: any) => mark?.attrs?.vertAlign === kind || mark?.type?.name === kind);
  }

  function currentStyleId(): string | null {
    try {
      const { $from } = editor.state.selection;
      for (let depth = $from.depth; depth >= 0; depth--) {
        const node = $from.node(depth);
        if (node.type.name === 'paragraph') return node.attrs?.paragraphProperties?.styleId ?? null;
      }
    } catch {
      // 选区不可用
    }
    return null;
  }

  function getSnapshot(): EngineSnapshot {
    const slice = (name: string): any => (ui ? safe(() => ui[name]?.getSnapshot?.() ?? null, null) : null);
    const state = { toolbar: slice('toolbar'), comments: slice('comments'), trackChanges: slice('trackChanges'), zoom: slice('zoom'), fonts: slice('fonts') };
    const commands: Partial<Record<DocxCommandId, CommandState>> = {};
    const raw = state?.toolbar?.commands ?? {};
    for (const [id, value] of Object.entries(raw as Record<string, CommandState | undefined>)) {
      if (value) commands[id as DocxCommandId] = { active: !!value.active, disabled: !!value.disabled, value: value.value };
    }
    const editable = !!editor?.isEditable;
    commands.superscript = { active: editor ? vertAlignActive('superscript') : false, disabled: !editable };
    commands.subscript = { active: editor ? vertAlignActive('subscript') : false, disabled: !editable };
    for (const id of ['cut', 'copy', 'paste', 'page-break', 'section-break', 'font-grow', 'font-shrink', 'multilevel-list'] as const) {
      commands[id] = { active: false, disabled: !editable };
    }
    const comments: EngineComment[] = (state?.comments?.items ?? []).map((item: any) => ({
      id: String(item.id),
      author: String(item.creatorName ?? item.author ?? ''),
      createdAt: isoOf(item.createdTime ?? item.date),
      text: String(item.text ?? ''),
      resolved: item.status === 'resolved' || item.isDone === true,
      parentId: item.parentCommentId ?? item.parentId ?? null,
      quote: String(item.anchoredText ?? '')
    }));
    const changes: EngineChange[] = (state?.trackChanges?.items ?? []).map((item: any) => {
      const change = item.change ?? item;
      const type = String(change.type ?? '');
      return {
        id: String(item.id),
        type: type === 'insert' || type === 'delete' || type === 'format' ? type : type === 'replacement' ? 'replace' : 'other',
        author: String(change.author ?? ''),
        createdAt: isoOf(change.date),
        excerpt: String(change.excerpt ?? change.insertedText ?? change.deletedText ?? ''),
        deletedText: change.deletedText ?? undefined,
        insertedText: change.insertedText ?? undefined
      } satisfies EngineChange;
    });
    const fonts = state?.fonts ?? {};
    return {
      ready: readyFlag,
      trackChanges,
      commands,
      styleId: editor ? currentStyleId() : null,
      comments,
      changes,
      activeCommentId: state?.comments?.activeIds?.[0] ?? null,
      zoom: Math.round(state?.zoom?.value ?? 100),
      pages,
      currentPage: Math.min(currentPage, pages),
      words,
      fontFamilies: (fonts.options ?? []).map((o: any) => ({ label: String(o.label ?? o.value ?? o), value: String(o.value ?? o.label ?? o) })),
      fontSizes: (fonts.sizeOptions ?? []).map((o: any) => ({ label: String(o.label), value: String(o.value) }))
    };
  }

  function safe<T>(fn: () => T, fallback: T): T {
    try {
      return fn();
    } catch (error) {
      log('warn', '文档编辑命令执行失败', { error: String(error) });
      return fallback;
    }
  }

  function selectionTarget(): unknown {
    return ui?.selection?.getSnapshot?.()?.selectionTarget ?? null;
  }

  function stepFontSize(direction: 1 | -1): boolean {
    const value = ui?.toolbar?.getSnapshot?.()?.commands?.['font-size']?.value;
    const current = ptOf(value) ?? 10.5;
    const next = direction > 0 ? SIZE_STEPS.find(s => s > current) ?? current : [...SIZE_STEPS].reverse().find(s => s < current) ?? current;
    return next !== current && ui.toolbar.execute('font-size', `${next}pt`);
  }

  function toggleVertAlign(kind: 'superscript' | 'subscript'): boolean {
    const target = selectionTarget();
    if (!target) return false;
    const value = vertAlignActive(kind) ? 'baseline' : kind;
    editor.doc.format.vertAlign({ target, value });
    return true;
  }

  async function paste(): Promise<boolean> {
    // 浏览器只允许读纯文本（需用户授权）；带格式粘贴请用 Ctrl+V
    const text = await navigator.clipboard?.readText?.();
    if (!text) return false;
    return !!editor.commands.insertContent?.(text);
  }

  function execute(id: DocxCommandId, payload?: unknown): boolean {
    if (!editor || !ui) return false;
    userTouched = true;
    return safe(() => {
      switch (id) {
        case 'superscript':
        case 'subscript':
          return toggleVertAlign(id);
        case 'font-grow':
          return stepFontSize(1);
        case 'font-shrink':
          return stepFontSize(-1);
        case 'cut':
        case 'copy':
          editor.view?.focus?.();
          return document.execCommand(id);
        case 'paste':
          editor.view?.focus?.();
          void paste().catch(error => log('warn', '读取剪贴板失败', { error: String(error) }));
          return true;
        case 'page-break':
          return !!editor.commands.insertPageBreak();
        case 'section-break':
          return !!editor.commands.insertSectionBreakAtSelection?.(payload ?? 'nextPage');
        case 'multilevel-list':
          // SuperDoc 1.x 没有「多级列表样式库」：先建编号列表，层级用「增加 / 减少缩进」（或 Tab）调整
          return ui.toolbar.execute('numbered-list');
        default:
          if (!BUILT_IN.has(id)) return false;
          return payload === undefined ? ui.toolbar.execute(id) : ui.toolbar.execute(id, payload);
      }
    }, false);
  }

  function setTrackChanges(on: boolean): void {
    userTouched = true;
    trackChanges = on;
    if (!readOnly) safe(() => superdoc.setDocumentMode(on ? 'suggesting' : 'editing'), undefined);
    log('info', on ? '修订已开启' : '修订已关闭');
    queueNotify();
  }

  function setReadOnly(on: boolean): void {
    readOnly = on;
    safe(() => superdoc.setDocumentMode(on ? 'viewing' : trackChanges ? 'suggesting' : 'editing'), undefined);
    queueNotify();
  }

  function styleMap(): Map<string, any> {
    const list: any[] = editor?.converter?.linkedStyles ?? [];
    return new Map(list.map(s => [s.id, s]));
  }

  function styles(): StyleOption[] {
    if (!editor) return [];
    const all = styleMap();
    const used = new Set<string>();
    editor.state.doc.descendants((node: any) => {
      if (node.type.name === 'paragraph') {
        const id = node.attrs?.paragraphProperties?.styleId;
        if (id) used.add(id);
        return false;
      }
      return true;
    });
    const options: StyleOption[] = [];
    for (const style of all.values()) {
      if (style.type !== 'paragraph') continue;
      const attrs = style.definition?.attrs ?? {};
      // 与 Word 样式库一致：快速样式（qFormat）或文档里实际用到的样式
      if (!attrs.qFormat && !used.has(style.id)) continue;
      const css = style.definition?.styles ?? {};
      const name = String(attrs.name ?? style.id);
      options.push({
        id: style.id,
        name,
        builtin: isBuiltinStyleName(name),
        headingLevel: headingLevelOf(style, all),
        fontSizePt: ptOf(css['font-size']),
        fontFamily: typeof css['font-family'] === 'string' ? css['font-family'] : undefined,
        bold: css['font-weight'] === 'bold' || css.bold === true
      });
    }
    // 正文在前，标题按级别，其余（含自定义）随后
    const rank = (o: StyleOption) => (/^normal$/i.test(o.name) ? 0 : o.headingLevel ? o.headingLevel : 20);
    return options.sort((a, b) => rank(a) - rank(b));
  }

  function applyStyle(styleId: string): boolean {
    userTouched = true;
    return safe(() => !!editor.commands.setStyleById(styleId), false);
  }

  function outline(): EngineOutlineItem[] {
    if (!editor) return [];
    const all = styleMap();
    const items: EngineOutlineItem[] = [];
    editor.state.doc.descendants((node: any) => {
      if (node.type.name !== 'paragraph') return true;
      const props = node.attrs?.paragraphProperties ?? {};
      const direct = Number.isInteger(props.outlineLvl) ? props.outlineLvl + 1 : 0;
      const level = direct || (props.styleId ? headingLevelOf(all.get(props.styleId), all) : 0);
      const text = String(node.textContent ?? '').trim();
      const id = node.attrs?.paraId ?? node.attrs?.sdBlockId;
      if (level >= 1 && level <= 3 && text && id) items.push({ level, text, id: String(id) });
      return false;
    });
    return items;
  }

  function scrollToBlock(blockId: string): void {
    void ui?.viewport
      ?.scrollIntoView({ target: { kind: 'text', blockId, range: { start: 0, end: 0 } }, block: 'start', behavior: 'smooth' })
      ?.catch?.(() => undefined);
  }

  function goToPage(page: number): void {
    const el = pageElements()[page - 1];
    el?.scrollIntoView({ block: 'start', behavior: 'smooth' });
  }

  function findResult(): FindResult {
    if (!search) return { total: 0, current: 0 };
    return { total: search.matches.length, current: search.matches.length ? search.index + 1 : 0 };
  }

  function runSearch(query: string, caseSensitive: boolean): any[] {
    const result = editor.commands.search(query, { caseSensitive });
    return Array.isArray(result) ? result : [];
  }

  function goToMatch(): void {
    const match = search?.matches[search.index];
    if (match) safe(() => editor.commands.goToSearchResult(match), undefined);
  }

  function find(query: string, opts: { caseSensitive?: boolean } = {}): FindResult {
    if (!editor || !query) {
      clearFind();
      return findResult();
    }
    const caseSensitive = !!opts.caseSensitive;
    const matches = safe(() => runSearch(query, caseSensitive), []);
    search = { query, caseSensitive, matches, index: matches.length ? 0 : -1 };
    goToMatch();
    return findResult();
  }

  function refreshSearch(): void {
    if (search && !search.matches.length) {
      const keep = Math.max(0, search.index);
      search.matches = safe(() => runSearch(search!.query, search!.caseSensitive), []);
      search.index = search.matches.length ? Math.min(keep, search.matches.length - 1) : -1;
    }
  }

  function step(delta: 1 | -1): FindResult {
    if (!search) return findResult();
    refreshSearch();
    if (!search.matches.length) return findResult();
    search.index = (search.index + delta + search.matches.length) % search.matches.length;
    goToMatch();
    return findResult();
  }

  function replaceCurrent(replacement: string): FindResult {
    userTouched = true;
    if (!search) return findResult();
    refreshSearch();
    const match = search.matches[search.index];
    if (!match) return findResult();
    safe(() => {
      editor.commands.goToSearchResult(match);
      editor.commands.insertContent(replacement);
    }, undefined);
    // 替换后重新搜，停在原序号
    const keep = search.index;
    search.matches = safe(() => runSearch(search!.query, search!.caseSensitive), []);
    search.index = search.matches.length ? Math.min(keep, search.matches.length - 1) : -1;
    goToMatch();
    return findResult();
  }

  function replaceAll(query: string, replacement: string): number {
    userTouched = true;
    if (!editor || !query) return 0;
    const matches = safe(() => runSearch(query, !!search?.caseSensitive), []);
    // 从后往前替换，前面的位置不受影响
    let count = 0;
    for (const match of [...matches].reverse()) {
      safe(() => {
        editor.commands.goToSearchResult(match);
        editor.commands.insertContent(replacement);
        count++;
      }, undefined);
    }
    clearFind();
    log('info', '全部替换', { count });
    return count;
  }

  function clearFind(): void {
    search = null;
    safe(() => editor?.commands?.clearSearchSession?.(), undefined);
  }

  function sectionTargets(): any[] {
    return safe(() => editor.doc.sections.list({}).items ?? [], []);
  }

  function getPageSetup(): PageSetup | null {
    const first = sectionTargets()[0];
    if (!first) return null;
    return {
      orientation: first.pageSetup?.orientation === 'landscape' ? 'landscape' : 'portrait',
      width: first.pageSetup?.width ?? 8.2681,
      height: first.pageSetup?.height ?? 11.6929,
      margins: {
        top: first.margins?.top ?? 1,
        right: first.margins?.right ?? 1,
        bottom: first.margins?.bottom ?? 1,
        left: first.margins?.left ?? 1
      }
    };
  }

  function setPageSetup(patch: Parameters<DocxEngine['setPageSetup']>[0]): boolean {
    userTouched = true;
    const items = sectionTargets();
    if (!items.length) return false;
    // 页面设置对全文各节一起生效（与 Word「应用于整篇文档」一致）
    return safe(() => {
      for (const item of items) {
        const landscape = (patch.orientation ?? item.pageSetup?.orientation) === 'landscape';
        if (patch.paper || patch.orientation) {
          const [w, h] = patch.paper ? PAPER[patch.paper]! : [Math.min(item.pageSetup.width, item.pageSetup.height), Math.max(item.pageSetup.width, item.pageSetup.height)];
          editor.doc.sections.setPageSetup({ target: item.address, orientation: landscape ? 'landscape' : 'portrait', width: landscape ? h : w, height: landscape ? w : h });
        }
        if (patch.margins) editor.doc.sections.setPageMargins({ target: item.address, ...patch.margins });
      }
      log('info', '页面设置已修改', { ...patch });
      return true;
    }, false);
  }

  const touch = <T>(value: T): T => {
    userTouched = true;
    return value;
  };
  const comments = {
    add: (text: string) => touch(safe(() => !!ui.comments.createFromSelection({ text }), false)),
    reply: (parentId: string, text: string) => touch(safe(() => !!ui.comments.reply(parentId, { text }), false)),
    resolve: (id: string) => touch(safe(() => !!ui.comments.resolve(id), false)),
    remove: (id: string) => touch(safe(() => !!ui.comments.delete(id), false)),
    focus: (id: string) => {
      safe(() => ui.comments.setActive(id), false);
      void safe(() => ui.comments.scrollTo(id), Promise.resolve())?.catch?.(() => undefined);
    }
  };

  function decideChange(id: string, accept: boolean): boolean {
    userTouched = true;
    const ok = safe(() => !!(accept ? ui.trackChanges.accept(id) : ui.trackChanges.reject(id)), false);
    if (ok) log('info', accept ? '接受修订' : '拒绝修订', { id });
    return ok;
  }

  async function exportDocx(): Promise<Blob> {
    if (!editor) throw new Error('docx editor not ready');
    const out = await editor.exportDocx({ isFinalDoc: false, commentsType: 'external' });
    return toBlob(out);
  }

  function summarize(): DocxSummary {
    return summarizeEditorDoc(editor.state.doc as PmNodeLike);
  }

  function destroy(): void {
    if (destroyed) return;
    destroyed = true;
    if (wordTimer) clearTimeout(wordTimer);
    cleanups.forEach(fn => fn());
    safe(() => ui?.destroy?.(), undefined);
    safe(() => superdoc.destroy?.(), undefined);
    host.remove();
    listeners.clear();
    contentListeners.clear();
  }

  return {
    ready,
    getSnapshot,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    onContentChange(listener) {
      contentListeners.add(listener);
      return () => contentListeners.delete(listener);
    },
    markSaved: rebaseline,
    execute,
    setTrackChanges,
    setReadOnly,
    styles,
    applyStyle,
    outline,
    goToOutline: scrollToBlock,
    goToPage,
    find,
    findNext: () => step(1),
    findPrevious: () => step(-1),
    replaceCurrent,
    replaceAll,
    clearFind,
    getPageSetup,
    setPageSetup,
    addComment: comments.add,
    replyComment: comments.reply,
    resolveComment: comments.resolve,
    deleteComment: comments.remove,
    focusComment: comments.focus,
    acceptChange: id => decideChange(id, true),
    rejectChange: id => decideChange(id, false),
    acceptAll: () => touch(safe(() => !!ui.trackChanges.acceptAll(), false)),
    rejectAll: () => touch(safe(() => !!ui.trackChanges.rejectAll(), false)),
    focusChange: id => {
      safe(() => ui.trackChanges.setActive(id), false);
      void safe(() => ui.trackChanges.scrollTo(id), Promise.resolve())?.catch?.(() => undefined);
    },
    exportDocx,
    summarize,
    destroy
  };
}
