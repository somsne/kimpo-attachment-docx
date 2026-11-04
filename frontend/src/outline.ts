/**
 * 导航窗格「标题」目录（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 查看态）。
 *
 * 按文档自带的标题样式抽目录：styles.xml 里每个段落样式的大纲级别（w:outlineLvl，或样式名为 heading 1～9，
 * 沿 basedOn 继承），段落自身 pPr 的 w:outlineLvl 优先；只取 1～3 级、文字非空的段落。
 * 中文 Word 文档的标题样式 ID 常是 "1" / "2"（样式名仍是 heading 1），所以按样式名与大纲级别判断，不按 ID。
 * 编辑态的目录由引擎从编辑器模型给出（同一口径）。
 */
import { attr, openDocxZip, readPart, scanTags, decodeXmlEntities } from './docx-xml.ts';

export interface OutlineItem {
  /** 1～3。 */
  level: number;
  text: string;
  /** 在全部目录项里的序号（从 0 起），查看态按它在渲染结果里找对应段落。 */
  index: number;
}

/** 最多取到几级标题。 */
export const OUTLINE_MAX_LEVEL = 3;

interface StyleInfo {
  name: string;
  basedOn: string | null;
  /** 0 起的大纲级别（w:outlineLvl 原值）；没有为 null。 */
  outline: number | null;
}

/** 解析 styles.xml：样式 ID → 名称、basedOn、大纲级别。 */
export function parseParagraphStyles(xml: string): Map<string, StyleInfo> {
  const styles = new Map<string, StyleInfo>();
  let current: { id: string; info: StyleInfo } | null = null;
  scanTags(xml, tag => {
    if (tag.name === 'w:style') {
      if (tag.closing) {
        if (current) styles.set(current.id, current.info);
        current = null;
      } else if (!tag.selfClosing && attr(tag.attrs, 'w:type') === 'paragraph') {
        const id = attr(tag.attrs, 'w:styleId');
        current = id ? { id, info: { name: '', basedOn: null, outline: null } } : null;
      }
      return;
    }
    if (!current || tag.closing) return;
    if (tag.name === 'w:name') current.info.name = attr(tag.attrs, 'w:val') ?? '';
    else if (tag.name === 'w:basedOn') current.info.basedOn = attr(tag.attrs, 'w:val');
    else if (tag.name === 'w:outlineLvl') {
      const level = Number(attr(tag.attrs, 'w:val'));
      if (Number.isInteger(level)) current.info.outline = level;
    }
  });
  return styles;
}

/** 样式的标题级别（1 起）；不是标题为 0。沿 basedOn 最多追 10 层。 */
export function headingLevelOfStyle(styles: Map<string, StyleInfo>, styleId: string | null): number {
  let id = styleId;
  for (let depth = 0; id && depth < 10; depth++) {
    const info = styles.get(id);
    if (!info) return 0;
    if (info.outline !== null) return info.outline >= 0 && info.outline < 9 ? info.outline + 1 : 0;
    const byName = /^heading\s*([1-9])$/i.exec(info.name.trim());
    if (byName) return Number(byName[1]);
    if (/^title$/i.test(info.name.trim())) return 1;
    id = info.basedOn;
  }
  return 0;
}

/** 从 document.xml 与 styles.xml 抽目录。 */
export function extractOutline(documentXml: string, stylesXml: string | null): OutlineItem[] {
  const styles = stylesXml ? parseParagraphStyles(stylesXml) : new Map<string, StyleInfo>();
  const items: OutlineItem[] = [];
  let para: { style: string | null; outline: number | null; text: string; depth: number } | null = null;
  let depth = 0;
  scanTags(documentXml, tag => {
    if (tag.name === 'w:p') {
      if (tag.selfClosing) return;
      if (!tag.closing) {
        depth++;
        // 只看最外层段落（文本框里的段落不进目录）
        if (depth === 1) para = { style: null, outline: null, text: '', depth };
      } else {
        if (depth === 1 && para) {
          const level = para.outline !== null ? (para.outline < 9 ? para.outline + 1 : 0) : headingLevelOfStyle(styles, para.style);
          const text = para.text.trim();
          if (level >= 1 && level <= OUTLINE_MAX_LEVEL && text) items.push({ level, text, index: items.length });
          para = null;
        }
        depth = Math.max(0, depth - 1);
      }
      return;
    }
    if (!para || tag.closing || depth !== 1) return;
    if (tag.name === 'w:pStyle') para.style = attr(tag.attrs, 'w:val');
    else if (tag.name === 'w:outlineLvl') {
      const level = Number(attr(tag.attrs, 'w:val'));
      if (Number.isInteger(level)) para.outline = level;
    } else if (tag.name === 'w:t' && !tag.selfClosing) {
      const close = documentXml.indexOf('</w:t>', tag.end);
      if (close < 0) return;
      para.text += decodeXmlEntities(documentXml.slice(tag.end, close));
      return close + 6;
    }
  });
  return items;
}

/** 从 docx 字节抽目录；失败得空目录（目录只是辅助，不影响查看）。 */
export async function outlineOfDocx(bytes: ArrayBuffer | Uint8Array): Promise<OutlineItem[]> {
  try {
    const zip = await openDocxZip(bytes);
    const documentXml = await readPart(zip, 'word/document.xml');
    if (!documentXml) return [];
    return extractOutline(documentXml, await readPart(zip, 'word/styles.xml'));
  } catch {
    return [];
  }
}

/**
 * 在 docx-preview 的渲染结果里按目录顺序找各标题对应的段落元素：只看正文（article 内）的 p，
 * 依次按「去空白文字相同」向后匹配；找不到的项得 null。
 */
export function locateOutlineElements(root: ParentNode, items: OutlineItem[]): Array<HTMLElement | null> {
  const paragraphs = Array.from(root.querySelectorAll<HTMLElement>('article p'));
  const norm = (s: string) => s.replace(/\s+/g, '');
  const found: Array<HTMLElement | null> = [];
  let from = 0;
  for (const item of items) {
    const want = norm(item.text);
    let hit: HTMLElement | null = null;
    for (let i = from; i < paragraphs.length; i++) {
      if (norm(paragraphs[i]!.textContent ?? '') === want) {
        hit = paragraphs[i]!;
        from = i + 1;
        break;
      }
    }
    found.push(hit);
  }
  return found;
}
