/**
 * 无损守门（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 打开时）。
 *
 * 方案文档「无损守门」：在线编辑引擎（SuperDoc）对一部分 OOXML 内容不能无损往返，含这些内容的文件只读打开，
 * 显示守门条幅（「请下载后编辑」），canEdit 在插件内降为假——宿主的权限判定不变，插件只是不进入编辑。
 * 扫描项（命中任一即只读）：
 *   - word/embeddings/      OLE 嵌入对象（Excel 表、Visio 图等）
 *   - word/diagrams/        SmartArt
 *   - word/charts/          图表（方案未单列，同属「在线编辑会丢失」，一并只读）
 *   - word/vbaProject.bin   宏
 *   - document.xml 里的 w:txbxContent（文本框）、mc:AlternateContent / wps:wsp（形状与文字混排）、
 *     w:object（OLE 对象占位）
 *   - document.xml 里的复杂域（w:fldChar 包起来的 w:instrText），域类型不在允许清单（TOC、MERGEFIELD，
 *     以及 TOC 条目自带的 PAGEREF / HYPERLINK 与页码 PAGE / NUMPAGES）里的
 *   - 打不开的包（不是 zip，如加密文档；或缺 word/document.xml）
 * 只读 document.xml 正文：页眉页脚里的页码域等不影响正文编辑。
 * 本文件只 import docx-xml.ts（带扩展名，供 Node 脚本直接加载）。
 */
import { openDocxZip, readPart, scanTags } from './docx-xml.ts';

export type GuardReason =
  | 'embeddings'
  | 'smartart'
  | 'charts'
  | 'macros'
  | 'textbox'
  | 'shapes'
  | 'ole_object'
  | 'complex_field'
  | 'unreadable';

export interface GuardResult {
  /** 命中任一项即只读。 */
  readOnly: boolean;
  reasons: GuardReason[];
  /** 命中的复杂域类型（大写，去重），供条幅与日志使用。 */
  fieldTypes: string[];
  /** settings.xml 里开着「修订」（w:trackRevisions），进入编辑时修订开关默认打开。 */
  trackRevisions: boolean;
}

/** settings.xml 是否开着修订（w:trackRevisions 且 w:val 不是 false / 0 / off）。 */
export function settingsTrackRevisions(xml: string | null): boolean {
  if (!xml) return false;
  const m = /<w:trackRevisions\b([^>]*)\/?>/.exec(xml);
  if (!m) return false;
  const val = /w:val\s*=\s*"([^"]*)"/.exec(m[1] ?? '')?.[1];
  return !(val && /^(false|0|off)$/i.test(val));
}

/** 允许在线编辑的复杂域类型。 */
export const ALLOWED_COMPLEX_FIELDS: ReadonlySet<string> = new Set(['TOC', 'MERGEFIELD', 'PAGEREF', 'HYPERLINK', 'PAGE', 'NUMPAGES']);

/** 扫描 document.xml：形状、文本框、OLE 占位、复杂域。 */
export function scanDocumentXml(xml: string): { reasons: Set<GuardReason>; fieldTypes: Set<string> } {
  const reasons = new Set<GuardReason>();
  const fieldTypes = new Set<string>();
  /** 复杂域栈：begin 入栈，instrText 拼到栈顶，separate（或没有 separate 时的 end）结算类型，end 出栈。 */
  const fields: Array<{ instr: string; done: boolean }> = [];
  const settle = (field: { instr: string; done: boolean } | undefined) => {
    if (!field || field.done) return;
    field.done = true;
    const type = field.instr.trim().split(/\s+/)[0]?.toUpperCase() ?? '';
    if (type && !ALLOWED_COMPLEX_FIELDS.has(type)) {
      reasons.add('complex_field');
      fieldTypes.add(type);
    }
  };
  scanTags(xml, tag => {
    if (tag.closing) return;
    switch (tag.name) {
      case 'w:txbxContent':
      case 'v:textbox':
        reasons.add('textbox');
        return;
      case 'mc:AlternateContent':
      case 'wps:wsp':
      case 'wpg:wgp':
        reasons.add('shapes');
        return;
      case 'w:object':
      case 'o:OLEObject':
        reasons.add('ole_object');
        return;
      case 'w:fldChar': {
        const kind = /w:fldCharType\s*=\s*"([^"]+)"/.exec(tag.attrs)?.[1];
        if (kind === 'begin') fields.push({ instr: '', done: false });
        else if (kind === 'separate') settle(fields[fields.length - 1]);
        else if (kind === 'end') settle(fields.pop());
        return;
      }
      case 'w:instrText': {
        if (tag.selfClosing) return;
        const close = xml.indexOf('</w:instrText>', tag.end);
        if (close < 0) return;
        const top = fields[fields.length - 1];
        if (top && !top.done) top.instr += xml.slice(tag.end, close);
        return close + 14;
      }
      default:
        return;
    }
  });
  return { reasons, fieldTypes };
}

/** 打开时扫描整个 docx 包。 */
export async function scanDocx(bytes: ArrayBuffer | Uint8Array): Promise<GuardResult> {
  let zip;
  try {
    zip = await openDocxZip(bytes);
  } catch {
    return { readOnly: true, reasons: ['unreadable'], fieldTypes: [], trackRevisions: false };
  }
  const reasons = new Set<GuardReason>();
  const names = Object.keys(zip.files);
  if (names.some(name => name.startsWith('word/embeddings/') && !zip.files[name]!.dir)) reasons.add('embeddings');
  if (names.some(name => name.startsWith('word/diagrams/') && !zip.files[name]!.dir)) reasons.add('smartart');
  if (names.some(name => name.startsWith('word/charts/') && !zip.files[name]!.dir)) reasons.add('charts');
  if (names.some(name => /^word\/vbaProject\.bin$/i.test(name))) reasons.add('macros');
  const xml = await readPart(zip, 'word/document.xml');
  let fieldTypes: string[] = [];
  if (xml == null) {
    reasons.add('unreadable');
  } else {
    const scanned = scanDocumentXml(xml);
    scanned.reasons.forEach(reason => reasons.add(reason));
    fieldTypes = [...scanned.fieldTypes].sort();
  }
  const list = [...reasons];
  const trackRevisions = settingsTrackRevisions(await readPart(zip, 'word/settings.xml'));
  return { readOnly: list.length > 0, reasons: list, fieldTypes, trackRevisions };
}
