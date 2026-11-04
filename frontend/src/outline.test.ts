// 导航目录：按样式名 / 大纲级别认标题（含中文文档常见的数字样式 ID）、basedOn 继承、只取 1～3 级；渲染结果里按文字定位。
import { describe, expect, it } from 'vitest';
import { extractOutline, headingLevelOfStyle, locateOutlineElements, outlineOfDocx, parseParagraphStyles } from './outline';
import { documentXml, makeDocx, para } from './test-helpers';

const styles = `<w:styles>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/></w:style>
  <w:style w:type="paragraph" w:styleId="ClauseTitle"><w:name w:val="条款标题"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="ClauseTitleBold"><w:name w:val="条款标题加粗"/><w:basedOn w:val="ClauseTitle"/></w:style>
  <w:style w:type="paragraph" w:styleId="H4"><w:name w:val="heading 4"/></w:style>
  <w:style w:type="character" w:styleId="1Char"><w:name w:val="标题 1 字符"/></w:style>
</w:styles>`;

describe('styles', () => {
  it('样式名 heading N、w:outlineLvl、basedOn 继承', () => {
    const map = parseParagraphStyles(styles);
    expect(headingLevelOfStyle(map, '1')).toBe(1);
    expect(headingLevelOfStyle(map, '2')).toBe(2);
    expect(headingLevelOfStyle(map, 'ClauseTitle')).toBe(3);
    expect(headingLevelOfStyle(map, 'ClauseTitleBold')).toBe(3);
    expect(headingLevelOfStyle(map, 'Normal')).toBe(0);
    expect(headingLevelOfStyle(map, 'missing')).toBe(0);
    expect(map.has('1Char')).toBe(false);
  });
});

describe('extractOutline', () => {
  it('抽 1～3 级非空标题，段落直设的 outlineLvl 优先，4 级与空标题不进目录', () => {
    const doc = documentXml(
      para('采购框架合同', '<w:pStyle w:val="1"/>') +
        para('正文', '<w:pStyle w:val="Normal"/>') +
        para('二、价格与付款', '<w:pStyle w:val="2"/>') +
        para('2.1 单价', '<w:pStyle w:val="ClauseTitle"/>') +
        para('四级', '<w:pStyle w:val="H4"/>') +
        para('', '<w:pStyle w:val="1"/>') +
        para('直设大纲', '<w:outlineLvl w:val="1"/>')
    );
    expect(extractOutline(doc, styles)).toEqual([
      { level: 1, text: '采购框架合同', index: 0 },
      { level: 2, text: '二、价格与付款', index: 1 },
      { level: 3, text: '2.1 单价', index: 2 },
      { level: 2, text: '直设大纲', index: 3 }
    ]);
  });

  it('outlineOfDocx 读包；坏包得空目录', async () => {
    const bytes = await makeDocx({ 'word/document.xml': documentXml(para('标题', '<w:pStyle w:val="1"/>')), 'word/styles.xml': styles });
    expect(await outlineOfDocx(bytes)).toHaveLength(1);
    expect(await outlineOfDocx(new Uint8Array([1, 2, 3]))).toEqual([]);
  });
});

describe('locateOutlineElements', () => {
  it('在 article 正文里按文字顺序找段落，页眉里的同名文字不算', () => {
    const root = document.createElement('div');
    root.innerHTML = '<section><header><p>采购框架合同</p></header><article><p>采购框架合同</p><p>正文</p><p>二、价格 与付款</p></article></section>';
    const found = locateOutlineElements(root, [
      { level: 1, text: '采购框架合同', index: 0 },
      { level: 2, text: '二、价格与付款', index: 1 },
      { level: 2, text: '不存在', index: 2 }
    ]);
    expect(found[0]?.parentElement?.tagName).toBe('ARTICLE');
    expect(found[1]?.textContent).toBe('二、价格 与付款');
    expect(found[2]).toBeNull();
  });
});
