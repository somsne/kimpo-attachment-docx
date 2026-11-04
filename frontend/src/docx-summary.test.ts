// 往返自检摘要：XML 口径（段落 / 表格 / 文字 / 图片，删除修订、AlternateContent 备份、合并续格不计）与比对结果。
import { describe, expect, it } from 'vitest';
import { compareSummaries, describeDiff, summarizeDocumentXml, summarizeDocx, summarizeEditorDoc, type PmNodeLike } from './docx-summary';
import { documentXml, makeDocx, para } from './test-helpers';

describe('summarizeDocumentXml', () => {
  it('段落、表格、文字与实体解码', () => {
    const xml = documentXml(
      para('第一段 &amp; &lt;符号&gt;') +
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
        para('末段')
    );
    const s = summarizeDocumentXml(xml);
    expect(s.paragraphs).toBe(4);
    expect(s.tables).toBe(1);
    expect(s.paragraphTexts).toEqual(['第一段 & <符号>', 'A1', '', '末段']);
    expect(s.text).toBe('第一段&<符号>A1末段');
  });

  it('删除修订（w:delText）不计、插入修订计入', () => {
    const xml = documentXml(
      '<w:p><w:r><w:t>单价</w:t></w:r><w:del w:id="1"><w:r><w:delText>不作调整</w:delText></w:r></w:del><w:ins w:id="2"><w:r><w:t>每年复核</w:t></w:r></w:ins></w:p>'
    );
    expect(summarizeDocumentXml(xml).text).toBe('单价每年复核');
  });

  it('mc:Fallback 里的备份内容不重复计入', () => {
    const xml = documentXml(
      '<w:p><mc:AlternateContent><mc:Choice><w:r><w:t>甲</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>甲</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>'
    );
    expect(summarizeDocumentXml(xml).text).toBe('甲');
  });

  it('纵向合并续格（vMerge 非 restart）里的段落不计', () => {
    const row = (first: string) => `<w:tr><w:tc>${first}</w:tc><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>`;
    const xml = documentXml(
      `<w:tbl>${row('<w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>合并</w:t></w:r></w:p>')}${row('<w:tcPr><w:vMerge/></w:tcPr><w:p/>')}</w:tbl>`
    );
    const s = summarizeDocumentXml(xml);
    expect(s.paragraphs).toBe(3);
    expect(s.text).toBe('合并xx');
  });

  it('图片计数（DrawingML pic:pic）', () => {
    const xml = documentXml('<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData><pic:pic/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>');
    expect(summarizeDocumentXml(xml).images).toBe(1);
  });

  it('summarizeDocx 读包内 word/document.xml；缺部件抛错', async () => {
    const bytes = await makeDocx({ 'word/document.xml': documentXml(para('你好')) });
    expect((await summarizeDocx(bytes)).text).toBe('你好');
    await expect(summarizeDocx(await makeDocx({}))).rejects.toThrow(/document\.xml/);
  });
});

describe('compareSummaries', () => {
  const a = summarizeDocumentXml(documentXml(para('甲') + para('乙')));

  it('只差空白算一致', () => {
    const b = summarizeDocumentXml(documentXml(para(' 甲 ') + para('乙')));
    expect(compareSummaries(a, b).ok).toBe(true);
  });

  it('文字不同 → 不一致，并指出第一处不同的段落', () => {
    const b = summarizeDocumentXml(documentXml(para('甲') + para('丙')));
    const diff = compareSummaries(a, b);
    expect(diff.ok).toBe(false);
    expect(diff.textEqual).toBe(false);
    expect(diff.firstTextDiff).toEqual({ index: 1, expected: '乙', actual: '丙' });
    expect(describeDiff(diff)).toContain('text differs');
  });

  it('段落数或表格数不同 → 不一致', () => {
    const b = summarizeDocumentXml(documentXml(para('甲乙')));
    expect(compareSummaries(a, b).ok).toBe(false);
    const withTable = summarizeDocumentXml(documentXml(para('甲') + para('乙') + '<w:tbl></w:tbl>'));
    const diff = compareSummaries(a, withTable);
    expect(diff.ok).toBe(false);
    expect(diff.tables).toEqual([0, 1]);
  });
});

describe('summarizeEditorDoc（编辑器模型口径）', () => {
  /** 造一个最小的 ProseMirror 形状节点。 */
  const node = (name: string, children: PmNodeLike[] = [], extra: Partial<PmNodeLike> = {}): PmNodeLike => ({
    type: { name },
    ...extra,
    descendants(cb) {
      const visit = (list: PmNodeLike[]) => {
        for (const child of list) if (cb(child) !== false) visit((child as unknown as { kids: PmNodeLike[] }).kids ?? []);
      };
      visit(children);
    },
    kids: children
  } as PmNodeLike);
  const text = (value: string, marks: string[] = []) => node('text', [], { isText: true, text: value, marks: marks.map(m => ({ type: { name: m } })) });

  it('段落 / 表格 / 图片计数，待接受的删除文字不计', () => {
    const doc = node('doc', [
      node('paragraph', [node('run', [text('单价'), text('不作调整', ['trackDelete']), text('每年复核', ['trackInsert'])])]),
      node('table', [node('tableRow', [node('tableCell', [node('paragraph', [text('A1')])])])]),
      node('paragraph', [node('image')])
    ]);
    const s = summarizeEditorDoc(doc);
    expect(s.paragraphs).toBe(3);
    expect(s.tables).toBe(1);
    expect(s.images).toBe(1);
    expect(s.text).toBe('单价每年复核A1');
  });
});
