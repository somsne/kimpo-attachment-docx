/**
 * OOXML（WordprocessingML）轻量扫描工具（文档附件插件 kimpo-attachment-docx，填报态附件外壳内）。
 *
 * 守门（guard.ts）、往返自检摘要（docx-summary.ts）、导航目录（outline.ts）共用。只做「按标签流扫描」，
 * 不建 DOM：浏览器与 Node（scripts/roundtrip.mjs 直接以类型剥离方式加载本文件）一套代码；
 * 本文件不 import 其他相对模块，便于脚本直接引用。
 */
import JSZip from 'jszip';

/** 一个标签事件：name 带命名空间前缀（如 w:p），attrs 为原样属性串。 */
export interface XmlTag {
  name: string;
  attrs: string;
  closing: boolean;
  selfClosing: boolean;
  /** 标签起点与终点（终点为 `>` 之后）在源串里的下标。 */
  start: number;
  end: number;
}

const TAG_RE = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;

/** 逐个产出标签（跳过注释、处理指令、CDATA 之外的普通标签）；visitor 返回数字时把扫描位置跳到该下标。 */
export function scanTags(xml: string, visitor: (tag: XmlTag) => number | void): void {
  TAG_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = TAG_RE.exec(xml))) {
    const tag: XmlTag = {
      name: match[2]!,
      attrs: match[3] ?? '',
      closing: match[1] === '/',
      selfClosing: match[4] === '/',
      start: match.index,
      end: match.index + match[0].length
    };
    const jump = visitor(tag);
    if (typeof jump === 'number' && jump > TAG_RE.lastIndex) TAG_RE.lastIndex = jump;
  }
}

/** 取属性值（如 attr(tag.attrs, 'w:val')）；没有得 null。 */
export function attr(attrs: string, name: string): string | null {
  const re = new RegExp(`(?:^|\\s)${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*("([^"]*)"|'([^']*)')`);
  const m = re.exec(attrs);
  return m ? decodeXmlEntities(m[2] ?? m[3] ?? '') : null;
}

/** XML 实体解码（五个预定义实体 + 数字实体）。 */
export function decodeXmlEntities(text: string): string {
  if (!text.includes('&')) return text;
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_all, ent: string) => {
    switch (ent) {
      case 'amp':
        return '&';
      case 'lt':
        return '<';
      case 'gt':
        return '>';
      case 'quot':
        return '"';
      case 'apos':
        return "'";
      default: {
        const code = ent[1] === 'x' ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
        try {
          return Number.isFinite(code) ? String.fromCodePoint(code) : '';
        } catch {
          return '';
        }
      }
    }
  });
}

/** 取 `<name …>…</name>` 的内文（从 openEnd 起找第一个同名闭合标签）；找不到得 null。 */
export function innerTextUntilClose(xml: string, openEnd: number, name: string): { text: string; next: number } | null {
  const close = xml.indexOf(`</${name}>`, openEnd);
  if (close < 0) return null;
  return { text: decodeXmlEntities(xml.slice(openEnd, close)), next: close + name.length + 3 };
}

/** 打开 docx（zip）；不是 zip（如加密 docx 是 OLE 复合文档）时抛错。 */
export async function openDocxZip(bytes: ArrayBuffer | Uint8Array): Promise<JSZip> {
  return JSZip.loadAsync(bytes);
}

/** 读包内文本部件；不存在得 null。 */
export async function readPart(zip: JSZip, path: string): Promise<string | null> {
  const file = zip.file(path);
  return file ? file.async('string') : null;
}

/** 去掉所有空白（往返比对口径：空白差异不算内容差异）。 */
export function stripWhitespace(text: string): string {
  return text.replace(/\s+/g, '');
}
