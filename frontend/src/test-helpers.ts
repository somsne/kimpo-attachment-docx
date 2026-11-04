// 单测公用：用 JSZip 现场拼最小 docx（守门、摘要、目录用例），不依赖二进制夹具。
import JSZip from 'jszip';

export const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

export function documentXml(body: string, extraNs = ''): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS} ${extraNs}><w:body>${body}<w:sectPr/></w:body></w:document>`;
}

export const para = (text: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

export async function makeDocx(parts: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  for (const [path, content] of Object.entries(parts)) zip.file(path, content);
  return zip.generateAsync({ type: 'uint8array' });
}
