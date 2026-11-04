/**
 * 往返自检的内容摘要（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 编辑态保存前）。
 *
 * 口径（方案文档「无损守门」）：段落数、表格数、正文文字（去空白）三项一致才放行保存；图片数一并统计，
 * 只用于差异报告与 roundtrip 脚本（不单独拦截：图片在 SuperDoc 里重排后锚定方式可能变化）。
 * 摘要只看 word/document.xml 的正文（w:body）：页眉页脚、脚注不在内；修订里的删除文字（w:delText）不计，
 * 插入文字计入；mc:AlternateContent 只取 Choice 分支（Fallback 是同一内容的备份，计入会重复）；
 * 纵向 / 横向合并单元格的续格（w:vMerge、w:hMerge 非 restart）不计——Word 只显示首格内容，编辑器合并后也不持有续格。
 *
 * 两个来源同一口径：
 *   - summarizeDocumentXml：从 docx 包解析（原文件、导出结果）；
 *   - SuperDoc 编辑器模型的摘要由引擎（src/engine）按同一口径从 ProseMirror 文档计算（见 summarizeEditorDoc）。
 * 本文件只 import docx-xml.ts（带扩展名，供 Node 脚本以类型剥离方式直接加载）。
 */
import { openDocxZip, readPart, scanTags, stripWhitespace, decodeXmlEntities } from './docx-xml.ts';

export interface DocxSummary {
  /** 正文段落数（含表格单元格里的段落）。 */
  paragraphs: number;
  /** 表格数（含嵌套表格）。 */
  tables: number;
  /** 图片数（w:drawing 内的 pic:pic 与 VML v:imagedata）。 */
  images: number;
  /** 每段文字（未去空白），按文档顺序。 */
  paragraphTexts: string[];
  /** 全部正文文字去空白后的拼接。 */
  text: string;
}

export interface SummaryDiff {
  ok: boolean;
  paragraphs: [number, number];
  tables: [number, number];
  images: [number, number];
  textEqual: boolean;
  /** 第一处文字不同的段落（按去空白文字逐段比对；段落数不同时从头比对到较短一侧结束）。 */
  firstTextDiff?: { index: number; expected: string; actual: string };
}

/** 从 document.xml 文本计算摘要。 */
export function summarizeDocumentXml(xml: string): DocxSummary {
  const bodyStart = xml.indexOf('<w:body');
  const bodyEnd = xml.lastIndexOf('</w:body>');
  const body = bodyStart >= 0 ? xml.slice(bodyStart, bodyEnd > bodyStart ? bodyEnd : undefined) : xml;
  const texts: string[] = [];
  /** 当前打开的段落在 texts 里的下标栈（文本框等会出现段落嵌套）。 */
  const stack: number[] = [];
  let tables = 0;
  let images = 0;
  let fallbackDepth = 0;
  /** 单元格栈：true = 合并续格（其中的段落与文字不计）。 */
  const cells: boolean[] = [];
  const inContinuation = () => cells.includes(true);
  scanTags(body, tag => {
    if (tag.name === 'mc:Fallback') {
      if (tag.selfClosing) return;
      fallbackDepth += tag.closing ? -1 : 1;
      return;
    }
    if (fallbackDepth > 0) return;
    switch (tag.name) {
      case 'w:tc':
        if (tag.closing) cells.pop();
        else if (!tag.selfClosing) cells.push(false);
        return;
      case 'w:vMerge':
      case 'w:hMerge': {
        if (tag.closing || !cells.length) return;
        const val = /w:val\s*=\s*"([^"]*)"/.exec(tag.attrs)?.[1];
        if (val !== 'restart') cells[cells.length - 1] = true;
        return;
      }
      case 'w:p':
        if (inContinuation()) return;
        if (tag.closing) stack.pop();
        else {
          texts.push('');
          if (!tag.selfClosing) stack.push(texts.length - 1);
        }
        return;
      case 'w:tbl':
        if (!tag.closing && !tag.selfClosing && !inContinuation()) tables++;
        return;
      case 'pic:pic':
      case 'v:imagedata':
        if (!tag.closing && !inContinuation()) images++;
        return;
      case 'w:t': {
        if (tag.closing || tag.selfClosing || inContinuation()) return;
        const close = body.indexOf('</w:t>', tag.end);
        if (close < 0) return;
        const at = stack[stack.length - 1];
        if (at !== undefined) texts[at] += decodeXmlEntities(body.slice(tag.end, close));
        return close + 6;
      }
      default:
        return;
    }
  });
  return { paragraphs: texts.length, tables, images, paragraphTexts: texts, text: stripWhitespace(texts.join('')) };
}

/** 从 docx 字节计算摘要；包里没有 word/document.xml 时抛错。 */
export async function summarizeDocx(bytes: ArrayBuffer | Uint8Array): Promise<DocxSummary> {
  const zip = await openDocxZip(bytes);
  const xml = await readPart(zip, 'word/document.xml');
  if (xml == null) throw new Error('word/document.xml missing');
  return summarizeDocumentXml(xml);
}

/** 比对两份摘要：段落数、表格数、正文文字（去空白）都一致才 ok。 */
export function compareSummaries(expected: DocxSummary, actual: DocxSummary): SummaryDiff {
  const textEqual = expected.text === actual.text;
  const diff: SummaryDiff = {
    ok: textEqual && expected.paragraphs === actual.paragraphs && expected.tables === actual.tables,
    paragraphs: [expected.paragraphs, actual.paragraphs],
    tables: [expected.tables, actual.tables],
    images: [expected.images, actual.images],
    textEqual
  };
  const n = Math.min(expected.paragraphTexts.length, actual.paragraphTexts.length);
  for (let i = 0; i < n; i++) {
    const a = stripWhitespace(expected.paragraphTexts[i]!);
    const b = stripWhitespace(actual.paragraphTexts[i]!);
    if (a !== b) {
      diff.firstTextDiff = { index: i, expected: expected.paragraphTexts[i]!, actual: actual.paragraphTexts[i]! };
      break;
    }
  }
  if (!diff.firstTextDiff && expected.paragraphTexts.length !== actual.paragraphTexts.length) {
    diff.firstTextDiff = { index: n, expected: expected.paragraphTexts[n] ?? '', actual: actual.paragraphTexts[n] ?? '' };
  }
  return diff;
}

/** ProseMirror 节点的最小形状（编辑器模型摘要用，不依赖 prosemirror 类型）。 */
export interface PmNodeLike {
  type: { name: string };
  isText?: boolean;
  text?: string | null;
  marks?: ReadonlyArray<{ type: { name: string } }>;
  descendants(callback: (node: PmNodeLike) => boolean | void): void;
}

/**
 * 编辑器（SuperDoc / ProseMirror）文档的摘要，口径同 summarizeDocumentXml：
 * paragraph 节点计段落、table 计表格、image 计图片；带 trackDelete 标记的文字（待接受的删除）不计。
 */
export function summarizeEditorDoc(doc: PmNodeLike): DocxSummary {
  const texts: string[] = [];
  let tables = 0;
  let images = 0;
  const walk = (node: PmNodeLike, current: number | null): void => {
    node.descendants(child => {
      const name = child.type.name;
      if (name === 'paragraph') {
        texts.push('');
        walk(child, texts.length - 1);
        return false;
      }
      if (name === 'table') tables++;
      else if (name === 'image') images++;
      if (child.isText && current !== null) {
        const deleted = child.marks?.some(mark => mark.type.name === 'trackDelete');
        if (!deleted) texts[current] += child.text ?? '';
      }
      return true;
    });
  };
  walk(doc, null);
  return { paragraphs: texts.length, tables, images, paragraphTexts: texts, text: stripWhitespace(texts.join('')) };
}

/** 摘要差异的简短描述（日志与脚本输出用）。 */
export function describeDiff(diff: SummaryDiff): string {
  const parts: string[] = [];
  if (diff.paragraphs[0] !== diff.paragraphs[1]) parts.push(`paragraphs ${diff.paragraphs[0]}→${diff.paragraphs[1]}`);
  if (diff.tables[0] !== diff.tables[1]) parts.push(`tables ${diff.tables[0]}→${diff.tables[1]}`);
  if (diff.images[0] !== diff.images[1]) parts.push(`images ${diff.images[0]}→${diff.images[1]}`);
  if (!diff.textEqual) parts.push('text differs');
  if (diff.firstTextDiff) parts.push(`first diff at paragraph #${diff.firstTextDiff.index}`);
  return parts.length ? parts.join('; ') : 'identical';
}
