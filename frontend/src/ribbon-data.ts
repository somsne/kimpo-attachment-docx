/**
 * 功能区静态数据与纯函数（文档附件插件 kimpo-attachment-docx，编辑态功能区）。
 * 中文字号名、调色板、行距档、页边距预设、样式显示名。
 */
import type { StyleOption } from './engine-types';
import type { Translate } from './i18n';

/** 中文字号名 → 磅值（Word 中文版口径）。 */
export const CJK_SIZES: Array<[string, number]> = [
  ['初号', 42], ['小初', 36], ['一号', 26], ['小一', 24], ['二号', 22], ['小二', 18], ['三号', 16], ['小三', 15],
  ['四号', 14], ['小四', 12], ['五号', 10.5], ['小五', 9], ['六号', 7.5], ['小六', 6.5], ['七号', 5.5], ['八号', 5]
];

/** 繁体中文界面的字号名。 */
const CJK_SIZES_TW: Record<string, string> = {
  初号: '初號', 小初: '小初', 一号: '一號', 小一: '小一', 二号: '二號', 小二: '小二', 三号: '三號', 小三: '小三',
  四号: '四號', 小四: '小四', 五号: '五號', 小五: '小五', 六号: '六號', 小六: '小六', 七号: '七號', 八号: '八號'
};

/** 阿拉伯数字字号档（磅）。 */
export const NUMERIC_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

export function isChineseLocale(locale: string): boolean {
  return /^zh\b/i.test(locale);
}

/** 字号下拉选项：中文界面先列中文字号名，再列数字。 */
export function fontSizeOptions(locale: string): Array<{ label: string; value: string }> {
  const numeric = NUMERIC_SIZES.map(pt => ({ label: String(pt), value: `${pt}pt` }));
  if (!isChineseLocale(locale)) return numeric;
  const tw = /^zh-(tw|hk|hant)/i.test(locale);
  return [...CJK_SIZES.map(([name, pt]) => ({ label: tw ? CJK_SIZES_TW[name] ?? name : name, value: `${pt}pt` })), ...numeric];
}

/** 磅值的显示名：中文界面能对上中文字号名时用字号名，否则数字。 */
export function fontSizeLabel(value: unknown, locale: string): string {
  const pt = parseFloat(String(value ?? ''));
  if (!Number.isFinite(pt) || pt <= 0) return '';
  if (isChineseLocale(locale)) {
    const hit = CJK_SIZES.find(([, size]) => Math.abs(size - pt) < 0.01);
    if (hit) return /^zh-(tw|hk|hant)/i.test(locale) ? CJK_SIZES_TW[hit[0]] ?? hit[0] : hit[0];
  }
  return String(Math.round(pt * 10) / 10);
}

/** 常用中文字体（文档自带字体列表里没有时补上；浏览器用系统字体渲染）。 */
export const CJK_FONTS = ['宋体', '黑体', '楷体', '仿宋', '微软雅黑', '等线'];

/** 字体颜色调色板（null = 自动）。 */
export const TEXT_COLORS: Array<string | null> = [
  null, '#000000', '#7f7f7f', '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'
];

/** 突出显示颜色（null = 无颜色）。 */
export const HIGHLIGHT_COLORS: Array<string | null> = ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#0000ff', '#ff0000', '#000080', '#008080', '#008000', '#800080', '#808080', '#c0c0c0', null];

export const LINE_HEIGHTS = [1, 1.15, 1.5, 2, 2.5, 3];

/** 页边距预设（英寸）。 */
export const MARGIN_PRESETS = {
  normal: { top: 1, bottom: 1, left: 1.25, right: 1.25 },
  narrow: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 },
  moderate: { top: 1, bottom: 1, left: 0.75, right: 0.75 },
  wide: { top: 1, bottom: 1, left: 2, right: 2 }
} as const;

export const PAPER_SIZES = ['A4', 'A3', 'B5', 'Letter', 'Legal'] as const;

const BUILTIN_STYLE_KEYS: Array<[RegExp, string]> = [
  [/^normal$/i, 'docx.style.normal'],
  [/^title$/i, 'docx.style.title'],
  [/^subtitle$/i, 'docx.style.subtitle'],
  [/^quote$/i, 'docx.style.quote'],
  [/^intense quote$/i, 'docx.style.intense_quote'],
  [/^list paragraph$/i, 'docx.style.list_paragraph'],
  [/^no spacing$/i, 'docx.style.no_spacing'],
  [/^caption$/i, 'docx.style.caption'],
  [/^strong$/i, 'docx.style.strong'],
  [/^emphasis$/i, 'docx.style.emphasis']
];

/** 样式在样式库里的显示名：Word 预置样式翻译成界面语言，自定义样式用文档里的原名。 */
export function styleDisplayName(style: Pick<StyleOption, 'name'>, t: Translate): string {
  const name = style.name.trim();
  const heading = /^heading\s*([1-9])$/i.exec(name);
  if (heading) return t('docx.style.heading', { n: heading[1] });
  for (const [re, key] of BUILTIN_STYLE_KEYS) if (re.test(name)) return t(key);
  return name;
}

/** 样式卡片的副标题：自定义样式标「自定义」，其余给字体与字号。 */
export function styleSubtitle(style: StyleOption, t: Translate, locale: string): string {
  if (!style.builtin) return t('docx.style.custom');
  const size = style.fontSizePt ? fontSizeLabel(style.fontSizePt, locale) : '';
  const sizeText = size && !isChineseLocale(locale) ? t('docx.style.size_pt', { n: size }) : size;
  return [style.fontFamily?.split(',')[0]?.replace(/["']/g, '').trim(), sizeText].filter(Boolean).join(' ');
}
