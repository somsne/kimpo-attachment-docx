// 无损守门：每类不支持内容各一份最小样件；允许的域（TOC / MERGEFIELD / 目录自带的 PAGEREF、HYPERLINK）不拦；settings 修订开关。
import { describe, expect, it } from 'vitest';
import { scanDocumentXml, scanDocx, settingsTrackRevisions } from './guard';
import { documentXml, makeDocx, para } from './test-helpers';

const plain = documentXml(para('正文'));
const field = (instr: string) =>
  `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>结果</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

describe('scanDocx', () => {
  it('普通文档不拦', async () => {
    const result = await scanDocx(await makeDocx({ 'word/document.xml': plain }));
    expect(result).toEqual({ readOnly: false, reasons: [], fieldTypes: [], trackRevisions: false });
  });

  it.each([
    ['embeddings', { 'word/embeddings/oleObject1.bin': new Uint8Array([1]) }],
    ['smartart', { 'word/diagrams/data1.xml': '<dgm:dataModel/>' }],
    ['charts', { 'word/charts/chart1.xml': '<c:chartSpace/>' }],
    ['macros', { 'word/vbaProject.bin': new Uint8Array([1]) }]
  ] as const)('包内部件命中 %s → 只读', async (reason, parts) => {
    const result = await scanDocx(await makeDocx({ 'word/document.xml': plain, ...parts }));
    expect(result.readOnly).toBe(true);
    expect(result.reasons).toContain(reason);
  });

  it('不是 zip（如加密文档）→ unreadable 只读', async () => {
    const result = await scanDocx(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0]));
    expect(result).toMatchObject({ readOnly: true, reasons: ['unreadable'] });
  });

  it('settings.xml 开着修订 → trackRevisions', async () => {
    const result = await scanDocx(await makeDocx({ 'word/document.xml': plain, 'word/settings.xml': '<w:settings><w:trackRevisions/></w:settings>' }));
    expect(result.trackRevisions).toBe(true);
  });
});

describe('scanDocumentXml', () => {
  it('文本框 / 形状 / OLE 占位', () => {
    expect([...scanDocumentXml(documentXml('<w:p><w:r><w:txbxContent>' + para('框内') + '</w:txbxContent></w:r></w:p>')).reasons]).toContain('textbox');
    expect([...scanDocumentXml(documentXml('<w:p><w:r><mc:AlternateContent><mc:Choice><wps:wsp/></mc:Choice></mc:AlternateContent></w:r></w:p>')).reasons]).toContain('shapes');
    expect([...scanDocumentXml(documentXml('<w:p><w:r><w:object><o:OLEObject/></w:object></w:r></w:p>')).reasons]).toContain('ole_object');
  });

  it('允许的复杂域不拦：TOC、MERGEFIELD、目录条目里嵌套的 PAGEREF / HYPERLINK、页码', () => {
    const toc =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\o "1-3" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>' +
      field('HYPERLINK \\l "_Toc1"') +
      field('PAGEREF _Toc1 \\h') +
      '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const result = scanDocumentXml(documentXml(toc + field('MERGEFIELD 客户名称') + field('PAGE')));
    expect([...result.reasons]).toEqual([]);
  });

  it('其他复杂域（IF、DOCPROPERTY 等）→ complex_field，记录域类型；指令拆在多个 instrText 里也能认出', () => {
    const split =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DOC</w:instrText></w:r><w:r><w:instrText>PROPERTY Title</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const result = scanDocumentXml(documentXml(field('IF 1 = 1 "是" "否"') + split));
    expect([...result.reasons]).toEqual(['complex_field']);
    expect([...result.fieldTypes].sort()).toEqual(['DOCPROPERTY', 'IF']);
  });

  it('w:fldSimple（简单域）不算复杂域', () => {
    expect([...scanDocumentXml(documentXml('<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>2026</w:t></w:r></w:fldSimple></w:p>')).reasons]).toEqual([]);
  });
});

describe('settingsTrackRevisions', () => {
  it.each([
    ['<w:trackRevisions/>', true],
    ['<w:trackRevisions w:val="true"/>', true],
    ['<w:trackRevisions w:val="false"/>', false],
    ['<w:trackRevisions w:val="0"/>', false],
    ['<w:zoom/>', false]
  ])('%s → %s', (xml, expected) => {
    expect(settingsTrackRevisions(`<w:settings>${xml}</w:settings>`)).toBe(expected);
  });
});
