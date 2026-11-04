/**
 * 字数统计（文档附件插件 kimpo-attachment-docx，状态栏「N 字」）。
 *
 * 按 Word 中文版口径：中日韩文字逐字计数，其余按连续的字母数字串计为一个词；标点与空白不计。
 */
const CJK = /[㐀-䶿一-鿿豈-﫿぀-ヿㇰ-ㇿ가-힯]/g;
const WORD = /[A-Za-z0-9À-ɏ]+(?:['’.\-][A-Za-z0-9À-ɏ]+)*/g;

export function countWords(text: string): number {
  const cjk = text.match(CJK)?.length ?? 0;
  const rest = text.replace(CJK, ' ');
  return cjk + (rest.match(WORD)?.length ?? 0);
}
