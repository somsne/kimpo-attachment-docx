<script setup lang="ts">
/**
 * 文档查看态（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 查看）。
 *
 * docx-preview（Apache-2.0）分页渲染：白纸页面（暗色下也是白纸），页面区底色取宿主主题变量。
 * 工具栏按外壳查看态画板（Main）口径：上一页 / 第 n / N 页 / 下一页 · 缩小 / 百分比 / 放大 / 适合宽度，
 * 左侧「导航窗格」开关，可下载时右侧出「下载」，最右「关于」。左侧导航窗格「标题 / 页面」：
 * 标题按文档标题样式抽取（outline.ts），点击跳到对应段落。缩放用 CSS zoom（Chromium 与 Firefox 126+）。
 * 渲染失败（损坏 / 加密 / 非 docx）时上报 error 并整块提示。
 */
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import Icon from './Icon.vue';
import { locateOutlineElements, type OutlineItem } from './outline';
import type { Translate } from './i18n';

const props = defineProps<{
  bytes: ArrayBuffer;
  outline: OutlineItem[];
  canDownload: boolean;
  t: Translate;
}>();

const emit = defineEmits<{
  (e: 'download'): void;
  (e: 'about'): void;
  (e: 'rendered', info: { pages: number }): void;
  (e: 'error', error: unknown): void;
}>();

/** docx-preview 的样式类名前缀（不用默认的 docx，避免与页面里别的 docx-preview 实例样式互相影响）。 */
const CLASS = 'kad-docx';
const ZOOM_STEPS = [50, 67, 75, 90, 100, 110, 125, 150, 175, 200];

const bodyRef = ref<HTMLElement | null>(null);
const styleRef = ref<HTMLElement | null>(null);
const scrollRef = ref<HTMLElement | null>(null);
const rendering = ref(true);
const failed = ref(false);
const pages = ref(1);
const page = ref(1);
const zoom = ref(100);
const navOpen = ref(true);
const navTab = ref<'headings' | 'pages'>('headings');
const activeHeading = ref(-1);
let headingEls: Array<HTMLElement | null> = [];
let renderSeq = 0;

const pageLabel = computed(() => props.t('docx.toolbar.page_of', { page: page.value, total: pages.value }));

function sections(): HTMLElement[] {
  return Array.from(bodyRef.value?.querySelectorAll<HTMLElement>(`section.${CLASS}`) ?? []);
}

async function render(): Promise<void> {
  const seq = ++renderSeq;
  rendering.value = true;
  failed.value = false;
  const body = bodyRef.value;
  const styles = styleRef.value;
  if (!body || !styles) return;
  try {
    // 按需引入：docx-preview 只在查看态用，单测里可整体替换
    const { renderAsync } = await import('docx-preview');
    body.innerHTML = '';
    styles.innerHTML = '';
    await renderAsync(props.bytes, body, styles, {
      className: CLASS,
      inWrapper: true,
      breakPages: true,
      // 用 Word 上次排版留下的分页标记，页数更接近桌面 Word
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      renderComments: false,
      renderChanges: false,
      experimental: false,
      useBase64URL: true
    });
    if (seq !== renderSeq) return;
    pages.value = Math.max(1, sections().length);
    page.value = 1;
    headingEls = locateOutlineElements(body, props.outline);
    rendering.value = false;
    emit('rendered', { pages: pages.value });
    await nextTick();
    updatePage();
  } catch (error) {
    if (seq !== renderSeq) return;
    rendering.value = false;
    failed.value = true;
    emit('error', error);
  }
}

/** 当前页：页面顶部越过滚动区上三分之一处的最后一页。 */
function updatePage(): void {
  const scroller = scrollRef.value;
  if (!scroller) return;
  const line = scroller.getBoundingClientRect().top + scroller.clientHeight / 3;
  let current = 1;
  sections().forEach((section, i) => {
    if (section.getBoundingClientRect().top <= line) current = i + 1;
  });
  page.value = current;
  // 导航窗格高亮当前所在的标题（刚点过目录时保持所点的项，直到用户自己滚动）
  if (pinnedHeading >= 0) {
    activeHeading.value = pinnedHeading;
    pinnedHeading = -1;
    return;
  }
  let active = -1;
  headingEls.forEach((el, i) => {
    if (el && el.getBoundingClientRect().top <= line) active = i;
  });
  activeHeading.value = active;
}

function goToPage(target: number): void {
  const list = sections();
  const section = list[Math.min(Math.max(target, 1), list.length) - 1];
  section?.scrollIntoView({ block: 'start' });
  page.value = Math.min(Math.max(target, 1), list.length || 1);
}

/** 点目录跳转后，滚动事件不覆盖刚选中的高亮（短文档里后面的标题也会越过判定线）。 */
let pinnedHeading = -1;

function goToHeading(index: number): void {
  const el = headingEls[index];
  if (!el) return;
  pinnedHeading = index;
  el.scrollIntoView({ block: 'start' });
  activeHeading.value = index;
}

function setZoom(value: number): void {
  zoom.value = Math.min(400, Math.max(25, Math.round(value)));
}

function zoomIn(): void {
  setZoom(ZOOM_STEPS.find(z => z > zoom.value) ?? zoom.value + 25);
}

function zoomOut(): void {
  setZoom([...ZOOM_STEPS].reverse().find(z => z < zoom.value) ?? zoom.value - 25);
}

/** 适合宽度：最宽一页加左右留白正好占满页面区。 */
function fitWidth(): void {
  const scroller = scrollRef.value;
  const widest = Math.max(0, ...sections().map(s => s.offsetWidth));
  if (!scroller || !widest) return;
  setZoom(((scroller.clientWidth - 48) / widest) * 100);
}

function onKeydown(event: KeyboardEvent): void {
  if (event.key === 'PageDown') {
    event.preventDefault();
    goToPage(page.value + 1);
  } else if (event.key === 'PageUp') {
    event.preventDefault();
    goToPage(page.value - 1);
  }
}

watch(() => props.bytes, () => void render());
watch(zoom, () => void nextTick(updatePage));

onMounted(() => {
  // 窄屏（外壳宽度不足）默认收起导航窗格
  const width = scrollRef.value?.parentElement?.clientWidth ?? 0;
  navOpen.value = !width || width >= 900;
  void render();
});

onBeforeUnmount(() => {
  renderSeq++;
});

defineExpose({ goToPage, goToHeading, setZoom, fitWidth, pages, page, zoom });
</script>

<template>
  <div class="kad-viewer" data-testid="docx-viewer">
    <div class="kad-toolbar" role="toolbar">
      <button type="button" class="kad-icon-btn" :class="{ 'is-on': navOpen }" :aria-pressed="navOpen" :aria-label="t('docx.toolbar.nav')" :title="t('docx.toolbar.nav')" data-testid="docx-nav-toggle" @click="navOpen = !navOpen"><Icon name="pane" /></button>
      <span class="kad-spacer" />
      <span class="kad-group">
        <button type="button" class="kad-icon-btn" :disabled="page <= 1" :aria-label="t('docx.toolbar.prev_page')" :title="t('docx.toolbar.prev_page')" data-testid="docx-prev" @click="goToPage(page - 1)"><Icon name="up" /></button>
        <span class="kad-label kad-page" data-testid="docx-page">{{ pageLabel }}</span>
        <button type="button" class="kad-icon-btn" :disabled="page >= pages" :aria-label="t('docx.toolbar.next_page')" :title="t('docx.toolbar.next_page')" data-testid="docx-next" @click="goToPage(page + 1)"><Icon name="down" /></button>
      </span>
      <span class="kad-divider" aria-hidden="true" />
      <span class="kad-group">
        <button type="button" class="kad-icon-btn" :aria-label="t('docx.toolbar.zoom_out')" :title="t('docx.toolbar.zoom_out')" data-testid="docx-zoom-out" @click="zoomOut">−</button>
        <span class="kad-label kad-zoom" data-testid="docx-zoom">{{ zoom }}%</span>
        <button type="button" class="kad-icon-btn" :aria-label="t('docx.toolbar.zoom_in')" :title="t('docx.toolbar.zoom_in')" data-testid="docx-zoom-in" @click="zoomIn">+</button>
        <button type="button" class="kad-text-btn" data-testid="docx-fit" @click="fitWidth">{{ t('docx.toolbar.fit_width') }}</button>
      </span>
      <span class="kad-spacer" />
      <button v-if="canDownload" type="button" class="kad-text-btn kad-with-icon" data-testid="docx-download" @click="emit('download')"><Icon name="download" :size="14" />{{ t('docx.toolbar.download') }}</button>
      <button type="button" class="kad-icon-btn" :aria-label="t('docx.toolbar.about')" :title="t('docx.toolbar.about')" data-testid="docx-about" @click="emit('about')"><Icon name="info" /></button>
    </div>

    <div class="kad-work">
      <nav v-if="navOpen" class="kad-nav" :aria-label="t('docx.toolbar.nav')" data-testid="docx-nav">
        <div class="kad-nav-tabs" role="tablist">
          <button type="button" role="tab" :aria-selected="navTab === 'headings'" :class="{ 'is-active': navTab === 'headings' }" @click="navTab = 'headings'">{{ t('docx.nav.headings') }}</button>
          <button type="button" role="tab" :aria-selected="navTab === 'pages'" :class="{ 'is-active': navTab === 'pages' }" data-testid="docx-nav-pages" @click="navTab = 'pages'">{{ t('docx.nav.pages') }}</button>
        </div>
        <div v-if="navTab === 'headings'" class="kad-nav-list" data-testid="docx-outline">
          <button
            v-for="item in outline"
            :key="item.index"
            type="button"
            class="kad-nav-item"
            :class="[`is-level-${item.level}`, { 'is-active': activeHeading === item.index }]"
            :title="item.text"
            @click="goToHeading(item.index)"
          >{{ item.text }}</button>
          <p v-if="!outline.length" class="kad-nav-empty">{{ t('docx.nav.no_headings') }}</p>
        </div>
        <div v-else class="kad-nav-list">
          <button v-for="n in pages" :key="n" type="button" class="kad-nav-item" :class="{ 'is-active': page === n }" @click="goToPage(n)">{{ t('docx.nav.page_item', { n }) }}</button>
        </div>
      </nav>

      <div ref="scrollRef" class="kad-canvas" tabindex="0" :aria-busy="rendering" @scroll.passive="updatePage" @keydown="onKeydown">
        <div ref="styleRef" class="kad-docx-styles" />
        <div ref="bodyRef" class="kad-docx-body" :style="{ zoom: zoom / 100 }" data-testid="docx-pages" />
        <div v-if="rendering" class="kad-overlay" role="status">
          <Icon name="spinner" :size="22" class="kad-spin" />
          <span>{{ t('docx.state.loading') }}</span>
        </div>
        <div v-if="failed" class="kad-overlay is-error" role="alert" data-testid="docx-render-failed">
          <Icon name="warning" :size="28" />
          <span>{{ t('docx.state.render_failed') }}</span>
          <button v-if="canDownload" type="button" class="kad-text-btn" @click="emit('download')">{{ t('docx.toolbar.download') }}</button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.kad-viewer {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
  flex-direction: column;
}

.kad-toolbar {
  flex: none;
  height: 44px;
  box-sizing: border-box;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 0 16px;
  border-bottom: 1px solid var(--kad-border);
  background: var(--kad-bg);
  font-size: 13px;
}

.kad-group {
  display: inline-flex;
  align-items: center;
  gap: 4px;
}

.kad-label {
  color: var(--kad-text-2);
  text-align: center;
  white-space: nowrap;
}

.kad-page {
  min-width: 110px;
}

.kad-zoom {
  min-width: 56px;
}

.kad-work {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
}

.kad-canvas {
  position: relative;
  flex: 1 1 auto;
  min-width: 0;
  overflow: auto;
  background: var(--kad-canvas);
  outline: none;
}

/* docx-preview 产出的外层与页面：页面区底色跟随主题，页面始终白纸 */
.kad-canvas :deep(.kad-docx-wrapper) {
  background: transparent;
  padding: 24px 24px 8px;
}

.kad-canvas :deep(.kad-docx-wrapper > section.kad-docx) {
  background: #ffffff;
  color: #1b1f24;
  box-shadow: var(--kad-page-shadow);
  margin-bottom: 24px;
}

.kad-overlay {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 12px;
  color: var(--kad-text-2);
  background: var(--kad-canvas);
  font-size: 14px;
}

.kad-overlay.is-error {
  color: var(--kad-text);
}
</style>
