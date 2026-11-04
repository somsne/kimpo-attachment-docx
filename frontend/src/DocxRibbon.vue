<script setup lang="ts">
/**
 * 编辑态功能区（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 编辑）。
 *
 * 按设计稿画板「文档插件 · 编辑态」：页签行 36px（开始 / 插入 / 布局 / 审阅 / 视图，右侧「修订」开关、
 * 查找替换、下载、关于），分组行 92px。「开始」五组严格照画板：剪贴板、字体、段落、样式库（文档自带样式含自定义）、
 * 插入（表格 / 图片）。其余页签沿用同一视觉口径（画板未逐项画出，按方案「编辑功能清单」补齐）。
 * 只经 DocxEngine 接口下命令，不碰 SuperDoc；SuperDoc 自带工具栏不显示。
 */
import { computed, onBeforeUnmount, onMounted, ref } from 'vue';
import Icon from './Icon.vue';
import type { IconName } from './icons';
import type { DocxCommandId, DocxEngine, EngineSnapshot, StyleOption } from './engine-types';
import type { Translate } from './i18n';
import {
  CJK_FONTS,
  HIGHLIGHT_COLORS,
  LINE_HEIGHTS,
  MARGIN_PRESETS,
  PAPER_SIZES,
  TEXT_COLORS,
  fontSizeLabel,
  fontSizeOptions,
  styleDisplayName,
  styleSubtitle
} from './ribbon-data';

export type RibbonTab = 'home' | 'insert' | 'layout' | 'review' | 'view';

const props = defineProps<{
  engine: DocxEngine;
  snapshot: EngineSnapshot;
  styles: StyleOption[];
  t: Translate;
  locale: string;
  /** 只读（守门、自检未过、失锁前的取锁中）时整排禁用。 */
  readOnly: boolean;
  canDownload: boolean;
  navOpen: boolean;
  reviewOpen: boolean;
  rulerOn: boolean;
}>();

const emit = defineEmits<{
  (e: 'find'): void;
  (e: 'download'): void;
  (e: 'about'): void;
  (e: 'new-comment'): void;
  (e: 'toggle-nav'): void;
  (e: 'toggle-review'): void;
  (e: 'toggle-ruler'): void;
}>();

const tab = ref<RibbonTab>('home');
const TABS: RibbonTab[] = ['home', 'insert', 'layout', 'review', 'view'];
/** 当前展开的下拉（同一时刻只开一个）。 */
const menu = ref<string | null>(null);
const gridHover = ref<[number, number]>([0, 0]);

const cmd = (id: DocxCommandId) => props.snapshot.commands[id];
const isActive = (id: DocxCommandId) => !!cmd(id)?.active;
const isDisabled = (id: DocxCommandId) => props.readOnly || !!cmd(id)?.disabled;

function run(id: DocxCommandId, payload?: unknown): void {
  menu.value = null;
  if (props.readOnly) return;
  props.engine.execute(id, payload);
}

function toggleMenu(name: string): void {
  menu.value = menu.value === name ? null : name;
}

/** 点浮层与下拉按钮以外的地方收起下拉（捕获期只读不拦，不影响编辑器自己的点击）。 */
function onDocMouseDown(event: MouseEvent): void {
  if (!menu.value) return;
  const target = event.target as Element | null;
  if (!target?.closest?.('.kad-pop, [data-menu]')) menu.value = null;
}

onMounted(() => document.addEventListener('mousedown', onDocMouseDown, true));
onBeforeUnmount(() => document.removeEventListener('mousedown', onDocMouseDown, true));

const fontFamily = computed(() => String(cmd('font-family')?.value ?? '').split(',')[0]?.replace(/["']/g, '').trim() || '');
const fontSize = computed(() => fontSizeLabel(cmd('font-size')?.value, props.locale));
const lineHeight = computed(() => {
  const v = Number(cmd('line-height')?.value);
  return Number.isFinite(v) && v > 0 ? String(Math.round(v * 100) / 100) : '1.0';
});
const align = computed(() => String(cmd('text-align')?.value ?? 'left'));
const textColor = computed(() => (cmd('text-color')?.value as string | null | undefined) ?? '#f5222d');
const highlightColor = computed(() => (cmd('highlight-color')?.value as string | null | undefined) ?? '#ffe58f');

const fontFamilies = computed(() => {
  const list = props.snapshot.fontFamilies.map(f => f.value);
  const names = [...new Set([...CJK_FONTS, ...list.map(v => v.split(',')[0]!.replace(/["']/g, '').trim())])].filter(Boolean);
  return names;
});
const sizeOptions = computed(() => fontSizeOptions(props.locale));

const galleryStyles = computed(() => props.styles.slice(0, 24));

interface SmallBtn {
  id: DocxCommandId;
  label: string;
  icon?: IconName;
  text?: string;
  textClass?: string;
  payload?: unknown;
  active?: boolean;
}

const fontButtons = computed<SmallBtn[]>(() => [
  { id: 'bold', label: props.t('docx.cmd.bold'), text: 'B', textClass: 'is-bold' },
  { id: 'italic', label: props.t('docx.cmd.italic'), text: 'I', textClass: 'is-italic' },
  { id: 'underline', label: props.t('docx.cmd.underline'), text: 'U', textClass: 'is-underline' },
  { id: 'strikethrough', label: props.t('docx.cmd.strike'), text: 'S', textClass: 'is-strike' },
  { id: 'superscript', label: props.t('docx.cmd.superscript'), text: 'x²', textClass: 'is-small' },
  { id: 'subscript', label: props.t('docx.cmd.subscript'), text: 'x₂', textClass: 'is-small' }
]);

const listButtons = computed<SmallBtn[]>(() => [
  { id: 'bullet-list', label: props.t('docx.cmd.bullets'), icon: 'bullets' },
  { id: 'numbered-list', label: props.t('docx.cmd.numbering'), icon: 'numbering' },
  { id: 'multilevel-list', label: props.t('docx.cmd.multilevel_hint'), icon: 'multilevel' },
  { id: 'indent-decrease', label: props.t('docx.cmd.indent_dec'), icon: 'indentDec' },
  { id: 'indent-increase', label: props.t('docx.cmd.indent_inc'), icon: 'indentInc' }
]);

const alignButtons = computed<SmallBtn[]>(() => [
  { id: 'text-align', payload: 'left', label: props.t('docx.cmd.align_left'), icon: 'alignLeft', active: align.value === 'left' },
  { id: 'text-align', payload: 'center', label: props.t('docx.cmd.align_center'), icon: 'alignCenter', active: align.value === 'center' },
  { id: 'text-align', payload: 'right', label: props.t('docx.cmd.align_right'), icon: 'alignRight', active: align.value === 'right' },
  { id: 'text-align', payload: 'justify', label: props.t('docx.cmd.align_justify'), icon: 'alignJustify', active: align.value === 'justify' || align.value === 'both' }
]);

const tableOps = computed<SmallBtn[]>(() => [
  { id: 'table-add-row-before', label: props.t('docx.cmd.row_above'), icon: 'rowAbove' },
  { id: 'table-add-row-after', label: props.t('docx.cmd.row_below'), icon: 'rowBelow' },
  { id: 'table-add-column-before', label: props.t('docx.cmd.col_left'), icon: 'colLeft' },
  { id: 'table-add-column-after', label: props.t('docx.cmd.col_right'), icon: 'colRight' },
  { id: 'table-delete-row', label: props.t('docx.cmd.delete_row'), icon: 'deleteRow' },
  { id: 'table-delete-column', label: props.t('docx.cmd.delete_col'), icon: 'deleteCol' },
  { id: 'table-merge-cells', label: props.t('docx.cmd.merge_cells'), icon: 'merge' },
  { id: 'table-split-cell', label: props.t('docx.cmd.split_cell'), icon: 'split' },
  { id: 'table-delete', label: props.t('docx.cmd.delete_table'), icon: 'trash' }
]);

function applyStyle(id: string): void {
  if (props.readOnly) return;
  props.engine.applyStyle(id);
}

function insertTable(rows: number, cols: number): void {
  run('table-insert', { rows, cols });
}

function insertLink(): void {
  menu.value = null;
  if (props.readOnly) return;
  // 链接地址用浏览器原生输入框（插件不自带对话框皮肤，宿主未提供输入框能力）
  const href = window.prompt(props.t('docx.cmd.link_prompt'), 'https://');
  if (href) props.engine.execute('link', { href });
}

function setPage(patch: Parameters<DocxEngine['setPageSetup']>[0]): void {
  menu.value = null;
  if (!props.readOnly) props.engine.setPageSetup(patch);
}

// 每次引擎快照更新（含改页面设置后）重新读一次
const pageSetup = computed(() => (props.snapshot && props.snapshot.ready ? props.engine.getPageSetup() : null));

function navigateChange(delta: 1 | -1): void {
  const list = props.snapshot.changes;
  if (!list.length) return;
  const at = list.findIndex(c => c.id === currentChange.value);
  const next = list[(at + delta + list.length) % list.length]!;
  currentChange.value = next.id;
  props.engine.focusChange(next.id);
}
const currentChange = ref<string | null>(null);

function toggleTrack(event: Event): void {
  props.engine.setTrackChanges((event.target as HTMLInputElement).checked);
}
</script>

<template>
  <div class="kad-ribbon" data-testid="docx-ribbon">
    <div class="kad-tabs" role="tablist">
      <button
        v-for="name in TABS"
        :key="name"
        type="button"
        role="tab"
        class="kad-tab"
        :class="{ 'is-active': tab === name }"
        :aria-selected="tab === name"
        :data-testid="`docx-tab-${name}`"
        @click="tab = name"
      >{{ t(`docx.tab.${name}`) }}</button>
      <span class="kad-spacer" />
      <label class="kad-track" data-testid="docx-track-toggle">
        <input type="checkbox" :checked="snapshot.trackChanges" :disabled="readOnly" @change="toggleTrack" />
        {{ t('docx.cmd.track_changes') }}
      </label>
      <button type="button" class="kad-tab-btn" data-testid="docx-find" @click="emit('find')"><Icon name="search" :size="14" />{{ t('docx.cmd.find_replace') }}</button>
      <button v-if="canDownload" type="button" class="kad-tab-btn" data-testid="docx-edit-download" @click="emit('download')"><Icon name="download" :size="14" />{{ t('docx.toolbar.download') }}</button>
      <button type="button" class="kad-tab-btn is-icon" :aria-label="t('docx.toolbar.about')" :title="t('docx.toolbar.about')" data-testid="docx-edit-about" @click="emit('about')"><Icon name="info" :size="14" /></button>
    </div>

    <div class="kad-groups" :class="{ 'is-readonly': readOnly }">
      <!-- 开始 -->
      <template v-if="tab === 'home'">
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :disabled="isDisabled('paste')" :title="t('docx.cmd.paste_hint')" data-testid="docx-cmd-paste" @click="run('paste')"><Icon name="paste" :size="22" />{{ t('docx.cmd.paste') }}</button>
            <div class="kad-col">
              <button type="button" class="kad-mini" :disabled="isDisabled('cut')" @click="run('cut')"><Icon name="cut" :size="13" />{{ t('docx.cmd.cut') }}</button>
              <button type="button" class="kad-mini" :disabled="isDisabled('copy')" @click="run('copy')"><Icon name="copy" :size="13" />{{ t('docx.cmd.copy') }}</button>
            </div>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.clipboard') }}</span>
        </div>

        <div class="kad-grp">
          <div class="kad-grp-body">
            <div class="kad-row">
              <div class="kad-dd-wrap">
                <button type="button" class="kad-select is-font" data-menu :disabled="readOnly" :aria-label="t('docx.cmd.font_family')" :title="t('docx.cmd.font_family')" data-testid="docx-font-family" @click="toggleMenu('font')"><span>{{ fontFamily || '—' }}</span><Icon name="caret" :size="12" /></button>
                <div v-if="menu === 'font'" class="kad-pop kad-scroll-pop">
                  <button v-for="name in fontFamilies" :key="name" type="button" class="kad-menu-item" :class="{ 'is-active': name === fontFamily }" :style="{ fontFamily: name }" @click="run('font-family', name)">{{ name }}</button>
                </div>
              </div>
              <div class="kad-dd-wrap">
                <button type="button" class="kad-select is-size" data-menu :disabled="readOnly" :aria-label="t('docx.cmd.font_size')" :title="t('docx.cmd.font_size')" data-testid="docx-font-size" @click="toggleMenu('size')"><span>{{ fontSize || '—' }}</span><Icon name="caret" :size="12" /></button>
                <div v-if="menu === 'size'" class="kad-pop kad-scroll-pop">
                  <button v-for="opt in sizeOptions" :key="opt.label" type="button" class="kad-menu-item" :class="{ 'is-active': opt.label === fontSize }" @click="run('font-size', opt.value)">{{ opt.label }}</button>
                </div>
              </div>
              <button type="button" class="kad-sq is-bordered" :disabled="isDisabled('font-grow')" :aria-label="t('docx.cmd.font_grow')" :title="t('docx.cmd.font_grow')" @click="run('font-grow')">A+</button>
              <button type="button" class="kad-sq is-bordered" :disabled="isDisabled('font-shrink')" :aria-label="t('docx.cmd.font_shrink')" :title="t('docx.cmd.font_shrink')" @click="run('font-shrink')">A−</button>
            </div>
            <div class="kad-row">
              <button
                v-for="b in fontButtons"
                :key="b.id"
                type="button"
                class="kad-sq"
                :class="[b.textClass, { 'is-on': isActive(b.id) }]"
                :aria-pressed="isActive(b.id)"
                :aria-label="b.label"
                :title="b.label"
                :disabled="isDisabled(b.id)"
                :data-testid="`docx-cmd-${b.id}`"
                @click="run(b.id)"
              >{{ b.text }}</button>
              <span class="kad-vsep" aria-hidden="true" />
              <div class="kad-dd-wrap">
                <button type="button" class="kad-sq is-color" data-menu :disabled="isDisabled('text-color')" :aria-label="t('docx.cmd.font_color')" :title="t('docx.cmd.font_color')" @click="toggleMenu('color')"><span class="kad-color-a">A</span><span class="kad-color-bar" :style="{ background: textColor ?? '#000' }" /></button>
                <div v-if="menu === 'color'" class="kad-pop kad-palette">
                  <button v-for="c in TEXT_COLORS" :key="String(c)" type="button" class="kad-swatch" :class="{ 'is-auto': !c }" :style="{ background: c ?? undefined }" :aria-label="c ?? t('docx.color.auto')" :title="c ?? t('docx.color.auto')" @click="run('text-color', c ?? 'auto')" />
                </div>
              </div>
              <div class="kad-dd-wrap">
                <button type="button" class="kad-sq is-color" data-menu :disabled="isDisabled('highlight-color')" :aria-label="t('docx.cmd.highlight')" :title="t('docx.cmd.highlight')" @click="toggleMenu('highlight')"><span class="kad-color-ab">ab</span><span class="kad-color-bar" :style="{ background: highlightColor ?? 'transparent' }" /></button>
                <div v-if="menu === 'highlight'" class="kad-pop kad-palette">
                  <button v-for="c in HIGHLIGHT_COLORS" :key="String(c)" type="button" class="kad-swatch" :class="{ 'is-auto': !c }" :style="{ background: c ?? undefined }" :aria-label="c ?? t('docx.color.none')" :title="c ?? t('docx.color.none')" @click="run('highlight-color', c ?? 'none')" />
                </div>
              </div>
              <button type="button" class="kad-sq is-small" :disabled="isDisabled('clear-formatting')" :aria-label="t('docx.cmd.clear_format')" :title="t('docx.cmd.clear_format')" @click="run('clear-formatting')">Aₓ</button>
            </div>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.font') }}</span>
        </div>

        <div class="kad-grp">
          <div class="kad-grp-body">
            <div class="kad-row">
              <button v-for="b in listButtons" :key="b.id" type="button" class="kad-sq" :class="{ 'is-on': isActive(b.id) }" :aria-pressed="isActive(b.id)" :aria-label="b.label" :title="b.label" :disabled="isDisabled(b.id)" :data-testid="`docx-cmd-${b.id}`" @click="run(b.id)"><Icon :name="b.icon!" /></button>
            </div>
            <div class="kad-row">
              <button v-for="b in alignButtons" :key="String(b.payload)" type="button" class="kad-sq" :class="{ 'is-on': b.active }" :aria-pressed="b.active" :aria-label="b.label" :title="b.label" :disabled="isDisabled('text-align')" :data-testid="`docx-align-${b.payload}`" @click="run('text-align', b.payload)"><Icon :name="b.icon!" /></button>
              <span class="kad-vsep" aria-hidden="true" />
              <div class="kad-dd-wrap">
                <button type="button" class="kad-select is-line" data-menu :disabled="isDisabled('line-height')" :aria-label="t('docx.cmd.line_spacing')" @click="toggleMenu('line')"><span>{{ t('docx.cmd.line_spacing_value', { v: lineHeight }) }}</span><Icon name="caret" :size="12" /></button>
                <div v-if="menu === 'line'" class="kad-pop">
                  <button v-for="v in LINE_HEIGHTS" :key="v" type="button" class="kad-menu-item" :class="{ 'is-active': String(v) === lineHeight }" @click="run('line-height', v)">{{ v.toFixed(v % 1 ? 2 : 1).replace(/0$/, '') }}</button>
                </div>
              </div>
              <button type="button" class="kad-sq" :class="{ 'is-on': isActive('formatting-marks') }" :aria-pressed="isActive('formatting-marks')" :aria-label="t('docx.cmd.marks')" :title="t('docx.cmd.marks')" :disabled="!!cmd('formatting-marks')?.disabled" @click="engine.execute('formatting-marks')">¶</button>
            </div>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.paragraph') }}</span>
        </div>

        <div class="kad-grp is-flex">
          <div class="kad-grp-body is-row kad-gallery" data-testid="docx-styles">
            <button
              v-for="s in galleryStyles"
              :key="s.id"
              type="button"
              class="kad-style"
              :class="{ 'is-on': snapshot.styleId === s.id || (!snapshot.styleId && /^normal$/i.test(s.name)) }"
              :disabled="readOnly"
              :title="styleDisplayName(s, t)"
              :data-style-id="s.id"
              @click="applyStyle(s.id)"
            >
              <span class="kad-style-name" :style="{ fontSize: s.headingLevel === 1 ? '15px' : s.headingLevel === 2 ? '14px' : '13px', fontWeight: s.headingLevel || s.bold ? 700 : 400 }">{{ styleDisplayName(s, t) }}</span>
              <span class="kad-style-sub">{{ styleSubtitle(s, t, locale) }}</span>
            </button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.styles') }}</span>
        </div>

        <div class="kad-grp is-last">
          <div class="kad-grp-body">
            <div class="kad-dd-wrap">
              <button type="button" class="kad-line-btn" data-menu :disabled="isDisabled('table-insert')" data-testid="docx-insert-table" @click="toggleMenu('table')"><Icon name="table" :size="14" />{{ t('docx.cmd.insert_table') }}</button>
              <div v-if="menu === 'table'" class="kad-pop kad-grid-pop" @mouseleave="gridHover = [0, 0]">
                <div class="kad-grid">
                  <template v-for="r in 8" :key="r">
                    <button v-for="c in 10" :key="`${r}-${c}`" type="button" class="kad-grid-cell" :class="{ 'is-on': r <= gridHover[0] && c <= gridHover[1] }" :aria-label="t('docx.cmd.table_size', { rows: r, cols: c })" @mouseenter="gridHover = [r, c]" @click="insertTable(r, c)" />
                  </template>
                </div>
                <div class="kad-grid-label">{{ gridHover[0] ? t('docx.cmd.table_size', { rows: gridHover[0], cols: gridHover[1] }) : t('docx.cmd.insert_table') }}</div>
              </div>
            </div>
            <button type="button" class="kad-line-btn" :disabled="isDisabled('image')" data-testid="docx-insert-image" @click="run('image')"><Icon name="image" :size="14" />{{ t('docx.cmd.insert_image') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.insert') }}</span>
        </div>
      </template>

      <!-- 插入 -->
      <template v-else-if="tab === 'insert'">
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <div class="kad-dd-wrap">
              <button type="button" class="kad-big" data-menu :disabled="isDisabled('table-insert')" @click="toggleMenu('table2')"><Icon name="table" :size="22" />{{ t('docx.cmd.insert_table') }}</button>
              <div v-if="menu === 'table2'" class="kad-pop kad-grid-pop" @mouseleave="gridHover = [0, 0]">
                <div class="kad-grid">
                  <template v-for="r in 8" :key="r">
                    <button v-for="c in 10" :key="`${r}-${c}`" type="button" class="kad-grid-cell" :class="{ 'is-on': r <= gridHover[0] && c <= gridHover[1] }" :aria-label="t('docx.cmd.table_size', { rows: r, cols: c })" @mouseenter="gridHover = [r, c]" @click="insertTable(r, c)" />
                  </template>
                </div>
                <div class="kad-grid-label">{{ gridHover[0] ? t('docx.cmd.table_size', { rows: gridHover[0], cols: gridHover[1] }) : t('docx.cmd.insert_table') }}</div>
              </div>
            </div>
            <div class="kad-ops">
              <button v-for="b in tableOps" :key="b.id" type="button" class="kad-sq" :aria-label="b.label" :title="b.label" :disabled="isDisabled(b.id)" @click="run(b.id)"><Icon :name="b.icon!" /></button>
            </div>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.tables') }}</span>
        </div>
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :disabled="isDisabled('image')" @click="run('image')"><Icon name="image" :size="22" />{{ t('docx.cmd.insert_image') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.illustrations') }}</span>
        </div>
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :disabled="isDisabled('page-break')" data-testid="docx-page-break" @click="run('page-break')"><Icon name="pageBreak" :size="22" />{{ t('docx.cmd.page_break') }}</button>
            <button type="button" class="kad-big" :disabled="isDisabled('table-of-contents-insert')" @click="run('table-of-contents-insert')"><Icon name="toc" :size="22" />{{ t('docx.cmd.toc') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.pages') }}</span>
        </div>
        <div class="kad-grp is-last">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :disabled="isDisabled('link')" @click="insertLink"><Icon name="link" :size="22" />{{ t('docx.cmd.link') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.links') }}</span>
        </div>
      </template>

      <!-- 布局 -->
      <template v-else-if="tab === 'layout'">
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <div class="kad-dd-wrap">
              <button type="button" class="kad-big" data-menu :disabled="readOnly" data-testid="docx-orientation" @click="toggleMenu('orient')"><Icon :name="pageSetup?.orientation === 'landscape' ? 'landscape' : 'orientation'" :size="22" />{{ t('docx.cmd.orientation') }}</button>
              <div v-if="menu === 'orient'" class="kad-pop">
                <button type="button" class="kad-menu-item" :class="{ 'is-active': pageSetup?.orientation !== 'landscape' }" @click="setPage({ orientation: 'portrait' })">{{ t('docx.cmd.portrait') }}</button>
                <button type="button" class="kad-menu-item" :class="{ 'is-active': pageSetup?.orientation === 'landscape' }" data-testid="docx-landscape" @click="setPage({ orientation: 'landscape' })">{{ t('docx.cmd.landscape') }}</button>
              </div>
            </div>
            <div class="kad-dd-wrap">
              <button type="button" class="kad-big" data-menu :disabled="readOnly" @click="toggleMenu('paper')"><Icon name="paper" :size="22" />{{ t('docx.cmd.paper') }}</button>
              <div v-if="menu === 'paper'" class="kad-pop">
                <button v-for="p in PAPER_SIZES" :key="p" type="button" class="kad-menu-item" @click="setPage({ paper: p })">{{ p }}</button>
              </div>
            </div>
            <div class="kad-dd-wrap">
              <button type="button" class="kad-big" data-menu :disabled="readOnly" @click="toggleMenu('margins')"><Icon name="margins" :size="22" />{{ t('docx.cmd.margins') }}</button>
              <div v-if="menu === 'margins'" class="kad-pop">
                <button v-for="(m, name) in MARGIN_PRESETS" :key="name" type="button" class="kad-menu-item" @click="setPage({ margins: { ...m } })">{{ t(`docx.cmd.margins_${name}`) }}</button>
              </div>
            </div>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.page_setup') }}</span>
        </div>
        <div class="kad-grp is-last">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :disabled="isDisabled('page-break')" @click="run('page-break')"><Icon name="pageBreak" :size="22" />{{ t('docx.cmd.page_break') }}</button>
            <div class="kad-col">
              <button type="button" class="kad-mini" :disabled="isDisabled('section-break')" @click="run('section-break', 'nextPage')"><Icon name="sectionBreak" :size="13" />{{ t('docx.cmd.section_next') }}</button>
              <button type="button" class="kad-mini" :disabled="isDisabled('section-break')" @click="run('section-break', 'continuous')"><Icon name="sectionBreak" :size="13" />{{ t('docx.cmd.section_continuous') }}</button>
            </div>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.breaks') }}</span>
        </div>
      </template>

      <!-- 审阅 -->
      <template v-else-if="tab === 'review'">
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :class="{ 'is-on': snapshot.trackChanges }" :aria-pressed="snapshot.trackChanges" :disabled="readOnly" data-testid="docx-review-track" @click="engine.setTrackChanges(!snapshot.trackChanges)"><Icon name="track" :size="22" />{{ t('docx.cmd.track_changes') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.tracking') }}</span>
        </div>
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :disabled="readOnly || isDisabled('track-changes-accept-selection')" data-testid="docx-accept-selection" @click="run('track-changes-accept-selection')"><Icon name="accept" :size="22" />{{ t('docx.cmd.accept') }}</button>
            <button type="button" class="kad-big" :disabled="readOnly || isDisabled('track-changes-reject-selection')" @click="run('track-changes-reject-selection')"><Icon name="reject" :size="22" />{{ t('docx.cmd.reject') }}</button>
            <div class="kad-col">
              <button type="button" class="kad-mini" :disabled="!snapshot.changes.length" @click="navigateChange(-1)"><Icon name="prev" :size="13" />{{ t('docx.cmd.prev_change') }}</button>
              <button type="button" class="kad-mini" :disabled="!snapshot.changes.length" @click="navigateChange(1)"><Icon name="next" :size="13" />{{ t('docx.cmd.next_change') }}</button>
            </div>
            <div class="kad-col">
              <button type="button" class="kad-mini" :disabled="readOnly || !snapshot.changes.length" data-testid="docx-accept-all" @click="engine.acceptAll()"><Icon name="accept" :size="13" />{{ t('docx.cmd.accept_all') }}</button>
              <button type="button" class="kad-mini" :disabled="readOnly || !snapshot.changes.length" data-testid="docx-reject-all" @click="engine.rejectAll()"><Icon name="reject" :size="13" />{{ t('docx.cmd.reject_all') }}</button>
            </div>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.changes') }}</span>
        </div>
        <div class="kad-grp is-last">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :disabled="readOnly" data-testid="docx-new-comment" @click="emit('new-comment')"><Icon name="comment" :size="22" />{{ t('docx.cmd.new_comment') }}</button>
            <button type="button" class="kad-big" :class="{ 'is-on': reviewOpen }" :aria-pressed="reviewOpen" @click="emit('toggle-review')"><Icon name="reviewPane" :size="22" />{{ t('docx.cmd.review_pane') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.comments') }}</span>
        </div>
      </template>

      <!-- 视图 -->
      <template v-else>
        <div class="kad-grp">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" :class="{ 'is-on': navOpen }" :aria-pressed="navOpen" data-testid="docx-view-nav" @click="emit('toggle-nav')"><Icon name="pane" :size="22" />{{ t('docx.cmd.nav_pane') }}</button>
            <button type="button" class="kad-big" :class="{ 'is-on': reviewOpen }" :aria-pressed="reviewOpen" @click="emit('toggle-review')"><Icon name="reviewPane" :size="22" />{{ t('docx.cmd.review_pane') }}</button>
            <button type="button" class="kad-big" :class="{ 'is-on': rulerOn }" :aria-pressed="rulerOn" @click="emit('toggle-ruler')"><Icon name="ruler" :size="22" />{{ t('docx.cmd.ruler') }}</button>
            <button type="button" class="kad-big" :class="{ 'is-on': isActive('formatting-marks') }" :aria-pressed="isActive('formatting-marks')" @click="engine.execute('formatting-marks')"><Icon name="marks" :size="22" />{{ t('docx.cmd.marks') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.show') }}</span>
        </div>
        <div class="kad-grp is-last">
          <div class="kad-grp-body is-row">
            <button type="button" class="kad-big" @click="engine.execute('zoom', 100)"><Icon name="zoom" :size="22" />{{ t('docx.cmd.zoom_100') }}</button>
            <button type="button" class="kad-big" @click="engine.execute('zoom-fit-width')"><Icon name="fitWidth" :size="22" />{{ t('docx.cmd.fit_width') }}</button>
          </div>
          <span class="kad-grp-label">{{ t('docx.group.zoom') }}</span>
        </div>
      </template>
    </div>
  </div>
</template>

<style scoped>
.kad-ribbon {
  flex: none;
  background: var(--kad-bg);
  color: var(--kad-text);
}

/* 页签行（画板：36px，白底，选中页签下划主色 2px） */
.kad-tabs {
  height: 36px;
  box-sizing: border-box;
  display: flex;
  align-items: flex-end;
  gap: 4px;
  padding: 0 16px;
  border-bottom: 1px solid var(--kad-border);
}

.kad-tab {
  padding: 8px 14px;
  border: 0;
  border-bottom: 2px solid transparent;
  background: transparent;
  color: var(--kad-text-2);
  font-size: 13px;
  cursor: pointer;
}

.kad-tab.is-active {
  border-bottom-color: var(--kad-primary);
  color: var(--kad-primary);
  font-weight: 600;
}

.kad-track {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 10px;
  color: var(--kad-text);
  cursor: pointer;
}

.kad-track input {
  margin: 0;
  accent-color: var(--kad-primary);
}

.kad-tab-btn {
  height: 28px;
  margin: 0 0 4px 6px;
  padding: 0 10px;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  border: 1px solid var(--kad-border);
  border-radius: 6px;
  background: var(--kad-bg);
  color: var(--kad-text);
  font-size: 12px;
  cursor: pointer;
}

.kad-tab-btn.is-icon {
  width: 28px;
  padding: 0;
  justify-content: center;
}

.kad-tab-btn:hover {
  border-color: var(--kad-primary);
  color: var(--kad-primary);
}

/* 分组行（画板：92px，分组间 1px 竖线，组名 11px 居中在底部） */
.kad-groups {
  height: 92px;
  box-sizing: border-box;
  display: flex;
  align-items: stretch;
  padding: 8px 16px;
  border-bottom: 1px solid var(--kad-border);
  /* 不设 overflow：下拉浮层要能伸出分组行 */
}

.kad-grp {
  flex: none;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  padding: 0 14px;
  border-right: 1px solid var(--kad-border);
}

.kad-grp:first-child {
  padding-left: 0;
}

.kad-grp.is-last {
  border-right: 0;
  padding-right: 0;
}

.kad-grp.is-flex {
  flex: 1 1 auto;
  min-width: 0;
}

.kad-grp-body {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.kad-grp-body.is-row {
  flex-direction: row;
  align-items: flex-start;
}

.kad-grp-label {
  font-size: 11px;
  color: var(--kad-text-2);
  text-align: center;
  white-space: nowrap;
}

.kad-row {
  display: flex;
  align-items: center;
  gap: 4px;
}

.kad-col {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.kad-ops {
  display: grid;
  grid-template-columns: repeat(5, 26px);
  gap: 2px 4px;
}

.kad-vsep {
  width: 1px;
  align-self: stretch;
  margin: 2px 4px;
  background: var(--kad-border);
}

.kad-big {
  min-width: 52px;
  height: 48px;
  padding: 0 6px;
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 4px;
  border: 0;
  border-radius: 6px;
  background: transparent;
  color: var(--kad-text);
  font-size: 11px;
  white-space: nowrap;
  cursor: pointer;
}

.kad-mini,
.kad-line-btn {
  height: 22px;
  padding: 0 8px;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  border: 0;
  border-radius: 4px;
  background: transparent;
  color: var(--kad-text);
  font-size: 11px;
  text-align: left;
  white-space: nowrap;
  cursor: pointer;
}

.kad-line-btn {
  height: 24px;
  font-size: 12px;
}

.kad-sq {
  width: 26px;
  height: 26px;
  padding: 0;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 0;
  border-radius: 4px;
  background: transparent;
  color: var(--kad-text);
  font-size: 13px;
  cursor: pointer;
}

.kad-sq.is-bordered {
  border: 1px solid var(--kad-border);
  background: var(--kad-bg);
  font-size: 12px;
}

.kad-sq.is-bold {
  font-weight: 700;
}

.kad-sq.is-italic {
  font-style: italic;
}

.kad-sq.is-underline {
  text-decoration: underline;
}

.kad-sq.is-strike {
  text-decoration: line-through;
}

.kad-sq.is-small {
  font-size: 11px;
}

.kad-sq.is-color {
  width: 30px;
  flex-direction: column;
  gap: 2px;
}

.kad-color-a {
  font-weight: 700;
  line-height: 1;
}

.kad-color-ab {
  font-size: 12px;
  line-height: 1;
}

.kad-color-bar {
  width: 18px;
  height: 3px;
}

.kad-big:hover:not(:disabled),
.kad-mini:hover:not(:disabled),
.kad-line-btn:hover:not(:disabled),
.kad-sq:hover:not(:disabled) {
  background: var(--kad-layout);
}

.kad-big.is-on,
.kad-sq.is-on {
  background: var(--kad-primary-soft);
  color: var(--kad-primary);
}

.kad-select {
  height: 26px;
  padding: 0 8px;
  display: inline-flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
  border: 1px solid var(--kad-border);
  border-radius: 4px;
  background: var(--kad-bg);
  color: var(--kad-text);
  font-size: 12px;
  cursor: pointer;
}

.kad-select span {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.kad-select.is-font {
  width: 120px;
}

.kad-select.is-size {
  width: 56px;
}

.kad-dd-wrap {
  position: relative;
}

.kad-scroll-pop {
  max-height: 260px;
  overflow: auto;
}

.kad-palette {
  display: grid;
  grid-template-columns: repeat(7, 20px);
  gap: 6px;
  min-width: 0;
}

.kad-swatch {
  width: 20px;
  height: 20px;
  padding: 0;
  border: 1px solid var(--kad-border);
  border-radius: 3px;
  cursor: pointer;
}

.kad-swatch.is-auto {
  background: linear-gradient(135deg, transparent 45%, var(--kad-error) 45%, var(--kad-error) 55%, transparent 55%), var(--kad-bg);
}

.kad-grid-pop {
  min-width: 0;
}

.kad-grid {
  display: grid;
  grid-template-columns: repeat(10, 16px);
  gap: 3px;
}

.kad-grid-cell {
  width: 16px;
  height: 16px;
  padding: 0;
  border: 1px solid var(--kad-border);
  border-radius: 2px;
  background: var(--kad-bg);
  cursor: pointer;
}

.kad-grid-cell.is-on {
  border-color: var(--kad-primary);
  background: var(--kad-primary-soft);
}

.kad-grid-label {
  margin-top: 6px;
  font-size: 12px;
  color: var(--kad-text-2);
  text-align: center;
}

/* 样式库（画板：86×56 卡片，当前样式描主色边并淡底） */
.kad-gallery {
  gap: 6px;
  overflow: hidden;
}

.kad-style {
  flex: none;
  width: 86px;
  height: 56px;
  padding: 6px 8px;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: flex-start;
  gap: 2px;
  border: 1px solid var(--kad-border);
  border-radius: 4px;
  background: var(--kad-bg);
  color: var(--kad-text);
  cursor: pointer;
  text-align: left;
}

.kad-style.is-on {
  border-color: var(--kad-primary);
  background: var(--kad-primary-faint);
}

.kad-style-name,
.kad-style-sub {
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.kad-style-sub {
  font-size: 10px;
  color: var(--kad-text-2);
}
</style>
