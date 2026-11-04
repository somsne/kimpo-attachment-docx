<script setup lang="ts">
/**
 * 文档编辑态（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 编辑，已取得编辑锁之后才挂载）。
 *
 * 布局按设计稿画板「文档插件 · 编辑态」：功能区（DocxRibbon）→ 工作区（左导航窗格 220px「标题 / 页面 / 批注」、
 * 中间页面区（SuperDoc 画的分页页面与标尺，页面白纸）、右审阅栏 240px）→ 状态栏 28px（第 n 页 / 共 N 页 · 字数 ·
 * 语言 · 修订开关状态 · 页面视图 · 缩放滑块与百分比）。
 * 挂载后按需加载编辑引擎（engine-loader.ts，随包 superdoc/engine.js），只经 DocxEngine 接口驱动。
 * 查找替换为浮动面板（Ctrl+F / Ctrl+H 打开，不拦截编辑器其余快捷键）。
 */
import { computed, onBeforeUnmount, onMounted, ref, shallowRef } from 'vue';
import Icon from './Icon.vue';
import DocxRibbon from './DocxRibbon.vue';
import ReviewPane from './ReviewPane.vue';
import type { DocxCommandId, DocxEngine, DocxEngineModule, EngineOutlineItem, EngineSnapshot, StyleOption } from './engine-types';
import { loadDocxEngine } from './engine-loader';
import type { Translate } from './i18n';

const props = defineProps<{
  bytes: ArrayBuffer;
  fileName: string;
  t: Translate;
  locale: string;
  timezone: string;
  userName: string;
  /** 进入编辑时修订开关的初值（文档 settings 里开着修订时为 true）。 */
  trackChanges: boolean;
  readOnly: boolean;
  canDownload: boolean;
  log: (level: 'info' | 'warn' | 'error', message: string, data?: Record<string, unknown>) => void;
  /** 单测注入假引擎；缺省加载随包引擎。 */
  loadEngine?: () => Promise<DocxEngineModule>;
}>();

const emit = defineEmits<{
  (e: 'ready', engine: DocxEngine): void;
  (e: 'change', modified: boolean): void;
  (e: 'error', error: unknown): void;
  (e: 'download'): void;
  (e: 'about'): void;
}>();

const EMPTY: EngineSnapshot = {
  ready: false,
  trackChanges: false,
  commands: {},
  styleId: null,
  comments: [],
  changes: [],
  activeCommentId: null,
  zoom: 100,
  pages: 1,
  currentPage: 1,
  words: 0,
  fontFamilies: [],
  fontSizes: []
};

const hostRef = ref<HTMLElement | null>(null);
const engine = shallowRef<DocxEngine | null>(null);
const snapshot = shallowRef<EngineSnapshot>(EMPTY);
const styleList = shallowRef<StyleOption[]>([]);
const outline = shallowRef<EngineOutlineItem[]>([]);
const loading = ref(true);
const failed = ref(false);
const navOpen = ref(true);
const navTab = ref<'headings' | 'pages' | 'comments'>('headings');
const reviewOpen = ref(true);
const rulerOn = ref(true);
const composing = ref(false);
const findOpen = ref(false);
const findQuery = ref('');
const replaceText = ref('');
const findCase = ref(false);
const findResult = ref<{ total: number; current: number } | null>(null);
const findNote = ref('');
const findInput = ref<HTMLInputElement | null>(null);
let offSnapshot: (() => void) | null = null;
let offContent: (() => void) | null = null;
let outlineTimer: ReturnType<typeof setTimeout> | null = null;
let disposed = false;

const rootComments = computed(() => snapshot.value.comments.filter(c => !c.parentId));

/** 插件自绘的右键菜单（SuperDoc 自带的是英文界面，已关闭）。位置相对页面区。 */
const menu = ref<{ x: number; y: number } | null>(null);
const pagesRef = ref<HTMLElement | null>(null);

interface MenuItem {
  key: string;
  label: string;
  run: () => void;
  disabled?: boolean;
  divider?: boolean;
}

const menuItems = computed<MenuItem[]>(() => {
  const e = engine.value;
  if (!e) return [];
  const cmd = (id: DocxCommandId) => snapshot.value.commands[id];
  const exec = (id: DocxCommandId) => () => e.execute(id);
  const items: MenuItem[] = [
    { key: 'cut', label: props.t('docx.cmd.cut'), run: exec('cut'), disabled: props.readOnly },
    { key: 'copy', label: props.t('docx.cmd.copy'), run: exec('copy') },
    { key: 'paste', label: props.t('docx.cmd.paste'), run: exec('paste'), disabled: props.readOnly },
    { key: 'comment', label: props.t('docx.cmd.new_comment'), run: newComment, disabled: props.readOnly, divider: true }
  ];
  const inTable = cmd('table-delete-row') && !cmd('table-delete-row')!.disabled;
  if (inTable && !props.readOnly) {
    const tableOps: Array<[DocxCommandId, string]> = [
      ['table-add-row-before', 'docx.cmd.row_above'],
      ['table-add-row-after', 'docx.cmd.row_below'],
      ['table-add-column-before', 'docx.cmd.col_left'],
      ['table-add-column-after', 'docx.cmd.col_right'],
      ['table-delete-row', 'docx.cmd.delete_row'],
      ['table-delete-column', 'docx.cmd.delete_col'],
      ['table-merge-cells', 'docx.cmd.merge_cells'],
      ['table-split-cell', 'docx.cmd.split_cell']
    ];
    tableOps.forEach(([id, key], i) => items.push({ key: id, label: props.t(key), run: exec(id), disabled: !!cmd(id)?.disabled, divider: i === 0 }));
  }
  const onChange = cmd('track-changes-accept-selection') && !cmd('track-changes-accept-selection')!.disabled;
  if (onChange && !props.readOnly) {
    items.push({ key: 'accept', label: props.t('docx.cmd.accept'), run: exec('track-changes-accept-selection'), divider: true });
    items.push({ key: 'reject', label: props.t('docx.cmd.reject'), run: exec('track-changes-reject-selection') });
  }
  return items;
});

function onContextMenu(event: MouseEvent): void {
  if (!engine.value || loading.value) return;
  const box = pagesRef.value?.getBoundingClientRect();
  if (!box) return;
  event.preventDefault();
  // 菜单不出页面区：靠右 / 靠下时向内收
  const x = Math.min(event.clientX - box.left, box.width - 200);
  const y = Math.min(event.clientY - box.top, box.height - 320);
  menu.value = { x: Math.max(4, x), y: Math.max(4, y) };
}

function runMenu(item: MenuItem): void {
  menu.value = null;
  if (!item.disabled) item.run();
}

function closeMenu(event: MouseEvent): void {
  if (menu.value && !(event.target as Element | null)?.closest?.('.kad-context-menu')) menu.value = null;
}

async function start(): Promise<void> {
  loading.value = true;
  failed.value = false;
  try {
    const mod = await (props.loadEngine ?? loadDocxEngine)();
    if (disposed || !hostRef.value) return;
    const created = mod.createDocxEngine({
      element: hostRef.value,
      file: new Blob([props.bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
      fileName: props.fileName,
      user: { name: props.userName },
      trackChanges: props.trackChanges,
      log: props.log
    });
    engine.value = created;
    offSnapshot = created.subscribe(next => {
      snapshot.value = next;
    });
    await created.ready;
    if (disposed) return;
    snapshot.value = created.getSnapshot();
    styleList.value = created.styles();
    outline.value = created.outline();
    offContent = created.onContentChange(modified => {
      emit('change', modified);
      // 目录与样式库随内容变化刷新（防抖）
      if (outlineTimer) clearTimeout(outlineTimer);
      outlineTimer = setTimeout(() => {
        outlineTimer = null;
        if (!engine.value) return;
        outline.value = engine.value.outline();
        styleList.value = engine.value.styles();
      }, 400);
    });
    loading.value = false;
    props.log('info', '文档编辑器已加载', { pages: snapshot.value.pages, styles: styleList.value.length });
    emit('ready', created);
  } catch (error) {
    if (disposed) return;
    loading.value = false;
    failed.value = true;
    props.log('error', '文档编辑器加载失败', { error: String(error) });
    emit('error', error);
  }
}

function retry(): void {
  teardown();
  disposed = false;
  void start();
}

function openFind(withReplace = false): void {
  findOpen.value = true;
  findNote.value = '';
  void Promise.resolve().then(() => findInput.value?.focus());
  if (withReplace) replaceText.value = replaceText.value || '';
}

function doFind(): void {
  if (!engine.value) return;
  findNote.value = '';
  findResult.value = engine.value.find(findQuery.value, { caseSensitive: findCase.value });
}

function step(delta: 1 | -1): void {
  if (!engine.value) return;
  if (!findResult.value) return doFind();
  findResult.value = delta > 0 ? engine.value.findNext() : engine.value.findPrevious();
}

function replaceOne(): void {
  if (!engine.value || props.readOnly) return;
  if (!findResult.value) doFind();
  findResult.value = engine.value.replaceCurrent(replaceText.value);
}

function replaceAll(): void {
  if (!engine.value || props.readOnly || !findQuery.value) return;
  const n = engine.value.replaceAll(findQuery.value, replaceText.value);
  findResult.value = null;
  findNote.value = props.t('docx.find.replaced', { n });
}

function closeFind(): void {
  findOpen.value = false;
  findResult.value = null;
  engine.value?.clearFind();
}

const findLabel = computed(() => {
  if (findNote.value) return findNote.value;
  if (!findResult.value) return '';
  return findResult.value.total ? props.t('docx.find.result', findResult.value) : props.t('docx.find.none');
});

/** Ctrl+F / Ctrl+H 打开查找替换（只在编辑器区域内生效，其余快捷键原样交给编辑器与浏览器）。 */
function onKeydown(event: KeyboardEvent): void {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
  const key = event.key.toLowerCase();
  if (key === 'f' || key === 'h') {
    event.preventDefault();
    openFind(key === 'h');
  }
}

function toggleRuler(): void {
  rulerOn.value = !rulerOn.value;
  engine.value?.execute('ruler');
}

function newComment(): void {
  reviewOpen.value = true;
  composing.value = true;
}

function setZoom(event: Event): void {
  engine.value?.execute('zoom', Number((event.target as HTMLInputElement).value));
}

function teardown(): void {
  disposed = true;
  if (outlineTimer) clearTimeout(outlineTimer);
  offSnapshot?.();
  offContent?.();
  offSnapshot = offContent = null;
  engine.value?.destroy();
  engine.value = null;
  snapshot.value = EMPTY;
}

onMounted(() => {
  document.addEventListener('mousedown', closeMenu, true);
  // 外壳宽度不足时默认收起两侧窗格（量不到宽度时按宽屏处理）
  const width = hostRef.value?.closest('.kad-root')?.clientWidth ?? 0;
  navOpen.value = !width || width >= 1100;
  reviewOpen.value = navOpen.value;
  void start();
});

onBeforeUnmount(() => {
  document.removeEventListener('mousedown', closeMenu, true);
  teardown();
});

defineExpose({ engine, snapshot, openFind });
</script>

<template>
  <div class="kad-editor" data-testid="docx-editor" @keydown="onKeydown">
    <DocxRibbon
      v-if="engine && !loading"
      :engine="engine"
      :snapshot="snapshot"
      :styles="styleList"
      :t="t"
      :locale="locale"
      :read-only="readOnly"
      :can-download="canDownload"
      :nav-open="navOpen"
      :review-open="reviewOpen"
      :ruler-on="rulerOn"
      @find="openFind()"
      @download="emit('download')"
      @about="emit('about')"
      @new-comment="newComment"
      @toggle-nav="navOpen = !navOpen"
      @toggle-review="reviewOpen = !reviewOpen"
      @toggle-ruler="toggleRuler"
    />

    <div class="kad-work">
      <nav v-if="navOpen && engine && !loading" class="kad-nav" :aria-label="t('docx.toolbar.nav')" data-testid="docx-edit-nav">
        <div class="kad-nav-tabs" role="tablist">
          <button type="button" role="tab" :aria-selected="navTab === 'headings'" :class="{ 'is-active': navTab === 'headings' }" @click="navTab = 'headings'">{{ t('docx.nav.headings') }}</button>
          <button type="button" role="tab" :aria-selected="navTab === 'pages'" :class="{ 'is-active': navTab === 'pages' }" @click="navTab = 'pages'">{{ t('docx.nav.pages') }}</button>
          <button type="button" role="tab" :aria-selected="navTab === 'comments'" :class="{ 'is-active': navTab === 'comments' }" @click="navTab = 'comments'">{{ t('docx.nav.comments', { n: rootComments.length }) }}</button>
        </div>
        <div v-if="navTab === 'headings'" class="kad-nav-list" data-testid="docx-edit-outline">
          <button v-for="item in outline" :key="item.id" type="button" class="kad-nav-item" :class="`is-level-${item.level}`" :title="item.text" @click="engine.goToOutline(item.id)">{{ item.text }}</button>
          <p v-if="!outline.length" class="kad-nav-empty">{{ t('docx.nav.no_headings') }}</p>
        </div>
        <div v-else-if="navTab === 'pages'" class="kad-nav-list">
          <button v-for="n in snapshot.pages" :key="n" type="button" class="kad-nav-item" :class="{ 'is-active': snapshot.currentPage === n }" @click="engine.goToPage(n)">{{ t('docx.nav.page_item', { n }) }}</button>
        </div>
        <div v-else class="kad-nav-list">
          <button v-for="c in rootComments" :key="c.id" type="button" class="kad-nav-item" :class="{ 'is-active': snapshot.activeCommentId === c.id }" :title="c.text" @click="engine.focusComment(c.id)">{{ c.author ? `${c.author}：` : '' }}{{ c.text }}</button>
          <p v-if="!rootComments.length" class="kad-nav-empty">{{ t('docx.nav.no_comments') }}</p>
        </div>
      </nav>

      <div ref="pagesRef" class="kad-pages" :class="{ 'no-ruler': !rulerOn }" @contextmenu="onContextMenu" @keydown.esc="menu = null">
        <div ref="hostRef" class="kad-engine-host" data-testid="docx-engine-host" />
        <div v-if="loading" class="kad-overlay" role="status" data-testid="docx-engine-loading">
          <Icon name="spinner" :size="22" class="kad-spin" />
          <span>{{ t('docx.state.engine_loading') }}</span>
        </div>
        <div v-else-if="failed" class="kad-overlay" role="alert" data-testid="docx-engine-failed">
          <Icon name="warning" :size="28" />
          <span>{{ t('docx.state.engine_failed') }}</span>
          <button type="button" class="kad-text-btn" @click="retry">{{ t('docx.state.retry') }}</button>
        </div>

        <div v-if="menu" class="kad-pop kad-context-menu" role="menu" :style="{ left: `${menu.x}px`, top: `${menu.y}px` }" data-testid="docx-context-menu">
          <template v-for="item in menuItems" :key="item.key">
            <div v-if="item.divider" class="kad-menu-divider" role="separator" />
            <button type="button" role="menuitem" class="kad-menu-item" :disabled="item.disabled" :data-menu-key="item.key" @click="runMenu(item)">{{ item.label }}</button>
          </template>
        </div>

        <div v-if="findOpen" class="kad-find" role="dialog" :aria-label="t('docx.find.title')" data-testid="docx-find-panel" @keydown.esc="closeFind">
          <div class="kad-find-row">
            <input ref="findInput" v-model="findQuery" class="kad-input kad-grow" :placeholder="t('docx.find.find')" :aria-label="t('docx.find.find')" data-testid="docx-find-input" @keydown.enter.prevent="findResult ? step(1) : doFind()" @input="findResult = null" />
            <span class="kad-find-count" data-testid="docx-find-count">{{ findLabel }}</span>
            <button type="button" class="kad-icon-btn is-sm" :aria-label="t('docx.find.prev')" :title="t('docx.find.prev')" @click="step(-1)"><Icon name="up" :size="14" /></button>
            <button type="button" class="kad-icon-btn is-sm" :aria-label="t('docx.find.next')" :title="t('docx.find.next')" data-testid="docx-find-next" @click="step(1)"><Icon name="down" :size="14" /></button>
            <button type="button" class="kad-icon-btn is-sm" :aria-label="t('docx.find.close')" :title="t('docx.find.close')" @click="closeFind"><Icon name="close" :size="14" /></button>
          </div>
          <div class="kad-find-row">
            <input v-model="replaceText" class="kad-input kad-grow" :placeholder="t('docx.find.replace')" :aria-label="t('docx.find.replace')" :disabled="readOnly" data-testid="docx-replace-input" />
            <button type="button" class="kad-text-btn is-sm" :disabled="readOnly || !findQuery" @click="replaceOne">{{ t('docx.find.replace_one') }}</button>
            <button type="button" class="kad-text-btn is-sm" :disabled="readOnly || !findQuery" data-testid="docx-replace-all" @click="replaceAll">{{ t('docx.find.replace_all') }}</button>
          </div>
          <label class="kad-find-case"><input v-model="findCase" type="checkbox" @change="findResult = null" />{{ t('docx.find.case') }}</label>
        </div>
      </div>

      <ReviewPane
        v-if="reviewOpen && engine && !loading"
        v-model:composing="composing"
        :engine="engine"
        :snapshot="snapshot"
        :t="t"
        :timezone="timezone"
        :read-only="readOnly"
      />
    </div>

    <div class="kad-status" data-testid="docx-status">
      <span data-testid="docx-status-page">{{ t('docx.status.page', { page: snapshot.currentPage, total: snapshot.pages }) }}</span>
      <span data-testid="docx-status-words">{{ t('docx.status.words', { n: snapshot.words.toLocaleString(locale) }) }}</span>
      <span>{{ t('docx.status.lang') }}</span>
      <span data-testid="docx-status-track">{{ snapshot.trackChanges ? t('docx.status.track_on') : t('docx.status.track_off') }}</span>
      <span class="kad-spacer" />
      <span>{{ t('docx.status.print_view') }}</span>
      <span class="kad-zoom-ctl">
        <button type="button" class="kad-zoom-btn" :aria-label="t('docx.toolbar.zoom_out')" :disabled="!engine" @click="engine?.execute('zoom', Math.max(50, snapshot.zoom - 10))">−</button>
        <input type="range" class="kad-zoom-range" min="50" max="200" step="10" :value="snapshot.zoom" :disabled="!engine" :aria-label="t('docx.group.zoom')" @change="setZoom" />
        <button type="button" class="kad-zoom-btn" :aria-label="t('docx.toolbar.zoom_in')" :disabled="!engine" @click="engine?.execute('zoom', Math.min(200, snapshot.zoom + 10))">+</button>
        <span class="kad-zoom-pct" data-testid="docx-status-zoom">{{ snapshot.zoom }}%</span>
      </span>
    </div>
  </div>
</template>

<style scoped>
.kad-editor {
  position: relative;
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
  flex-direction: column;
}

.kad-work {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
}

/* 页面区（画板：#e9ebef 底，页面白纸带阴影） */
.kad-pages {
  position: relative;
  flex: 1 1 auto;
  min-width: 0;
  background: var(--kad-canvas);
  overflow: hidden;
}

.kad-engine-host {
  position: absolute;
  inset: 0;
}

.kad-overlay {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 12px;
  background: var(--kad-canvas);
  color: var(--kad-text-2);
  font-size: 14px;
}

/* SuperDoc 画出的页面：页面外底色跟随主题（画板 #e9ebef），页面白纸带阴影、居中；暗色下也保持白纸 */
.kad-pages :deep(.superdoc),
.kad-pages :deep(.superdoc__layers),
.kad-pages :deep(.super-editor-container),
.kad-pages :deep(.super-editor) {
  background: transparent;
}

.kad-pages :deep(.super-editor-container) {
  display: flex;
  flex-direction: column;
  align-items: center;
}

.kad-pages :deep(.superdoc__layers) {
  flex: 1 1 auto;
}

.kad-pages :deep(.superdoc-page) {
  background: #ffffff;
  color: #1b1f24;
  box-shadow: var(--kad-page-shadow);
}

/* 标尺（画板：22px 灰底条，刻度只在页面宽度内） */
.kad-pages :deep(.ruler-host) {
  position: sticky;
  top: 0;
  z-index: 2;
  align-self: stretch;
  display: flex;
  justify-content: center;
  background: var(--kad-layout);
  border-bottom: 1px solid var(--kad-border);
  margin-bottom: 16px;
}

.kad-pages.no-ruler :deep(.ruler-host) {
  display: none;
}

.kad-context-menu {
  top: auto;
  min-width: 180px;
}

.kad-menu-divider {
  height: 1px;
  margin: 4px 0;
  background: var(--kad-border);
}

.kad-find {
  position: absolute;
  top: 12px;
  right: 20px;
  z-index: 20;
  width: 360px;
  box-sizing: border-box;
  padding: 10px;
  display: flex;
  flex-direction: column;
  gap: 6px;
  background: var(--kad-bg);
  border: 1px solid var(--kad-border);
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
  font-size: 12px;
}

.kad-find-row {
  display: flex;
  align-items: center;
  gap: 4px;
}

.kad-find-count {
  min-width: 48px;
  color: var(--kad-text-2);
  text-align: center;
  white-space: nowrap;
}

.kad-icon-btn.is-sm {
  width: 26px;
  height: 26px;
}

.kad-text-btn.is-sm {
  height: 26px;
  padding: 0 8px;
  font-size: 12px;
}

.kad-find-case {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: var(--kad-text-2);
}

/* 状态栏（画板：28px，白底，12px 次要色） */
.kad-status {
  flex: none;
  height: 28px;
  box-sizing: border-box;
  display: flex;
  align-items: center;
  gap: 18px;
  padding: 0 16px;
  background: var(--kad-bg);
  border-top: 1px solid var(--kad-border);
  font-size: 12px;
  color: var(--kad-text-2);
  white-space: nowrap;
}

.kad-zoom-ctl {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.kad-zoom-btn {
  padding: 0 2px;
  border: 0;
  background: transparent;
  color: var(--kad-text-2);
  font-size: 14px;
  cursor: pointer;
}

.kad-zoom-range {
  width: 120px;
  accent-color: var(--kad-primary);
}

.kad-zoom-pct {
  min-width: 36px;
}
</style>
