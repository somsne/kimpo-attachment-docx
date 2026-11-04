// @vitest-environment node
// 往返夹具（fixtures/roundtrip/*.docx，由 scripts/make-fixtures.mjs 生成）：真实 SuperDoc（与插件编辑态同一版本）
// headless 导入 → 导出 → 重新解析，三方摘要（原文件 / 编辑器模型 / 导出结果）零差异。
// 另做一次反向注入：改掉导出结果里的一段文字，自检必须报不一致（防判据恒真）。
import { describe, expect, it } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
// @ts-expect-error -- 纯 JS 脚本模块，无类型声明
import { exportBytes, openHeadless, roundtripDocx } from '../scripts/roundtrip-core.mjs';
import { compareSummaries, describeDiff, summarizeDocx, summarizeEditorDoc } from './docx-summary';

const dir = join(dirname(fileURLToPath(import.meta.url)), '..', 'fixtures', 'roundtrip');
const fixtures = readdirSync(dir).filter(name => name.endsWith('.docx')).sort();

describe('往返夹具零差异', () => {
  it('夹具齐全（8 份）', () => {
    expect(fixtures).toHaveLength(8);
  });

  it.each(fixtures)('%s', async name => {
    const result = await roundtripDocx(readFileSync(join(dir, name)));
    expect(result.guard.readOnly, name).toBe(false);
    expect(describeDiff(result.importDiff), 'import').toBe('identical');
    expect(describeDiff(result.roundtripDiff), 'roundtrip').toBe('identical');
    expect(describeDiff(result.saveCheckDiff), 'save-check').toBe('identical');
    expect(result.exported.images, 'images').toBe(result.original.images);
  });
});

describe('编辑后的保存自检', () => {
  it('改字后导出：编辑器模型与导出结果一致；与原文件不一致', async () => {
    const original = readFileSync(join(dir, '01-headings.docx'));
    const editor = await openHeadless(original);
    try {
      editor.commands.insertContentAt(1, '新增');
      const expected = summarizeEditorDoc(editor.state.doc);
      const exported = await summarizeDocx(await exportBytes(editor));
      expect(compareSummaries(expected, exported).ok).toBe(true);
      expect(compareSummaries(await summarizeDocx(original), exported).ok).toBe(false);
      expect(exported.text).toContain('新增');
    } finally {
      editor.destroy();
    }
  });

  it('反向注入：导出结果被篡改一段文字 → 自检不通过', async () => {
    const editor = await openHeadless(readFileSync(join(dir, '03-tables.docx')));
    try {
      const expected = summarizeEditorDoc(editor.state.doc);
      const zip = await JSZip.loadAsync(await exportBytes(editor));
      const xml = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', xml.replace('付款节点', '付款期限'));
      const tampered = await summarizeDocx(await zip.generateAsync({ type: 'uint8array' }));
      const diff = compareSummaries(expected, tampered);
      expect(diff.ok).toBe(false);
      expect(diff.firstTextDiff?.actual).toContain('付款期限');
    } finally {
      editor.destroy();
    }
  });
});
