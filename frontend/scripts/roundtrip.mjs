#!/usr/bin/env node
// PoC 往返验收脚本（文档附件插件 kimpo-attachment-docx）：给一个目录，对其中每个 .docx 跑
// 导入（SuperDoc headless，与插件编辑态同一版本）→ 导出 → 再解析，比对段落数、表格数、正文文字（去空白）与图片数，
// 输出逐件结果与汇总。供主会话用客户样件验收（门槛：20 件文字零差异）。页数差需要排版引擎，本脚本不测。
//
// 用法：  node scripts/roundtrip.mjs <目录> [--out <导出目录>] [--json]
//   --out   把导出结果写到该目录（同名文件），便于在 Word 里人工比对版式与页数
//   --json  输出机器可读的 JSON（每件一条）
// 退出码：全部零差异为 0，有差异或打不开的为 1。
import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import { roundtripDocx } from './roundtrip-core.mjs';
import { describeDiff } from '../src/docx-summary.ts';

const args = process.argv.slice(2);
const dir = args.find(a => !a.startsWith('--'));
const outIdx = args.indexOf('--out');
const outDir = outIdx >= 0 ? args[outIdx + 1] : null;
const asJson = args.includes('--json');
if (!dir || !statSync(dir, { throwIfNoEntry: false })?.isDirectory()) {
  console.error('用法：node scripts/roundtrip.mjs <目录> [--out <导出目录>] [--json]');
  process.exit(2);
}
if (outDir) mkdirSync(outDir, { recursive: true });

const files = readdirSync(dir).filter(n => /\.docx$/i.test(n) && !n.startsWith('~$')).sort();
const rows = [];
for (const name of files) {
  const path = join(dir, name);
  const started = Date.now();
  try {
    const r = await roundtripDocx(readFileSync(path));
    if (outDir) writeFileSync(join(outDir, basename(name)), r.exportedBytes);
    rows.push({
      file: name,
      ok: r.roundtripDiff.ok && r.importDiff.ok,
      guard: r.guard.reasons,
      paragraphs: r.roundtripDiff.paragraphs,
      tables: r.roundtripDiff.tables,
      images: r.roundtripDiff.images,
      textEqual: r.roundtripDiff.textEqual,
      import: describeDiff(r.importDiff),
      roundtrip: describeDiff(r.roundtripDiff),
      firstDiff: r.roundtripDiff.firstTextDiff ?? r.importDiff.firstTextDiff ?? null,
      ms: Date.now() - started
    });
  } catch (error) {
    rows.push({ file: name, ok: false, error: String(error?.message ?? error), ms: Date.now() - started });
  }
}

if (asJson) {
  for (const row of rows) console.log(JSON.stringify(row));
} else {
  for (const row of rows) {
    if (row.error) {
      console.log(`✗ ${row.file}  打不开：${row.error}`);
      continue;
    }
    const guard = row.guard.length ? `  [守门只读：${row.guard.join(',')}]` : '';
    console.log(`${row.ok ? '✓' : '✗'} ${row.file}  段落 ${row.paragraphs.join('→')}  表格 ${row.tables.join('→')}  图片 ${row.images.join('→')}  文字${row.textEqual ? '一致' : '不同'}  ${row.ms} ms${guard}`);
    if (!row.ok) {
      console.log(`    导入：${row.import}；往返：${row.roundtrip}`);
      if (row.firstDiff) console.log(`    第 ${row.firstDiff.index} 段  原：${JSON.stringify(row.firstDiff.expected.slice(0, 80))}  现：${JSON.stringify(row.firstDiff.actual.slice(0, 80))}`);
    }
  }
  const ok = rows.filter(r => r.ok).length;
  const guarded = rows.filter(r => r.guard?.length).length;
  console.log(`\n共 ${rows.length} 件：零差异 ${ok}，有差异或打不开 ${rows.length - ok}（其中守门只读 ${guarded} 件，编辑态不会保存它们）。目录：${resolve(dir)}`);
}
process.exit(rows.every(r => r.ok) ? 0 : 1);
