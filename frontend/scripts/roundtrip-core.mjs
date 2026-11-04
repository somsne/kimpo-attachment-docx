// 往返核心（文档附件插件 kimpo-attachment-docx）：docx → SuperDoc（headless，同一编辑引擎与版本）导入 → 导出 →
// 重新解析，比对段落数、表格数、正文文字（去空白），图片数一并报告。scripts/roundtrip.mjs 与
// src/roundtrip.test.ts 共用。依赖 Node 22 的 TypeScript 类型剥离直接加载 src 下的摘要与守门模块。
import { JSDOM } from 'jsdom';
import { compareSummaries, summarizeDocx, summarizeEditorDoc } from '../src/docx-summary.ts';
import { scanDocx } from '../src/guard.ts';

let editorModule = null;
async function loadEditor() {
  editorModule ??= await import('superdoc/super-editor');
  return editorModule;
}

/** 用 headless SuperDoc 打开 docx 字节；调用方负责 destroy。 */
export async function openHeadless(bytes) {
  const { Editor, getStarterExtensions } = await loadEditor();
  const { window } = new JSDOM('<!doctype html><html><body></body></html>');
  const buffer = Buffer.from(bytes);
  const [content, media, mediaFiles, fonts] = await Editor.loadXmlData(buffer, true);
  return new Editor({
    mode: 'docx',
    isHeadless: true,
    document: window.document,
    extensions: getStarterExtensions(),
    content,
    media,
    mediaFiles,
    fonts,
    documentId: 'kimpo-roundtrip',
    // 与插件编辑态一致：关掉 SuperDoc 默认开启的遥测
    telemetry: { enabled: false }
  });
}

/** 导出为字节（headless 下 SuperDoc 返回 Buffer）。 */
export async function exportBytes(editor) {
  const out = await editor.exportDocx({ isFinalDoc: false, commentsType: 'external' });
  if (!out) throw new Error('export returned nothing');
  if (out instanceof Uint8Array) return out;
  if (typeof out.arrayBuffer === 'function') return new Uint8Array(await out.arrayBuffer());
  return new Uint8Array(out);
}

/**
 * 对一份 docx 跑往返：
 *   original  原文件摘要（XML）
 *   imported  导入后编辑器模型摘要（与原文件比 → 导入是否丢内容）
 *   exported  导出再解析的摘要（与原文件比 → 往返是否丢内容；与 imported 比 → 保存自检口径）
 */
export async function roundtripDocx(bytes) {
  const guard = await scanDocx(bytes);
  const original = await summarizeDocx(bytes);
  const editor = await openHeadless(bytes);
  try {
    const imported = summarizeEditorDoc(editor.state.doc);
    const exportedBytes = await exportBytes(editor);
    const exported = await summarizeDocx(exportedBytes);
    return {
      guard,
      original,
      imported,
      exported,
      exportedBytes,
      importDiff: compareSummaries(original, imported),
      roundtripDiff: compareSummaries(original, exported),
      saveCheckDiff: compareSummaries(imported, exported)
    };
  } finally {
    editor.destroy();
  }
}
