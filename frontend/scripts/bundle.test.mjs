// 产物断言（构建后运行：pnpm build && pnpm run test:bundle）：
//   - plugin.umd.js：挂到 window.KimpoPlugins.attachment_docx、导出 createSurface，vue 外部化，样式内联，不含 SuperDoc；
//   - superdoc/engine.js：IIFE 挂到 window.KimpoPlugins.attachment_docx_engine（版本 1.47.2），vue 外部化（用宿主 Vue），
//     遥测关闭写死在封装里，SuperDoc 内联 process 垫片的 NODE_ENV 判断已定为 production，Chromium 88 垫片在位；
//   - 随包许可：dist/LICENSE（本插件 AGPL-3.0）、dist/SOURCE.txt（源码地址）、dist/superdoc/LICENSE（SuperDoc AGPL-3.0）。
import assert from 'node:assert/strict';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import vm from 'node:vm';
import { Window } from 'happy-dom';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const dist = join(root, 'dist');
const require = createRequire(import.meta.url);
const bundle = readFileSync(join(dist, 'plugin.umd.js'), 'utf8');
const engine = readFileSync(join(dist, 'superdoc', 'engine.js'), 'utf8');

test('UMD 挂载到 KimpoPlugins.attachment_docx 并导出 createSurface', () => {
  const sandbox = { Vue: require('vue'), console };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(bundle, sandbox, { filename: 'plugin.umd.js' });
  const createSurface = sandbox.KimpoPlugins?.attachment_docx?.createSurface;
  assert.equal(typeof createSurface, 'function');
  const mounted = createSurface({ surfaceType: 'attachment.open', hostCaps: [], props: {}, hostBridge: { marker: 1 } });
  assert.ok(mounted.component);
  assert.equal(mounted.props.hostBridge.marker, 1);
});

test('主包：vue 外部化、样式内联、不含 SuperDoc 编辑引擎', () => {
  assert.doesNotMatch(bundle, /__VUE_OPTIONS_API__|createRenderer\(/, '不应打包 Vue 运行时');
  assert.match(bundle, /kimpo-plugin-style-attachment_docx/);
  assert.doesNotMatch(bundle, /ingest\.superdoc\.dev|SuperConverter/, '主包不应含 SuperDoc');
  assert.match(bundle, /superdoc\/engine\.js/, '主包按需加载随包引擎');
  assert.ok(statSync(join(dist, 'plugin.umd.js')).size < 1024 * 1024, '主包应小于 1 MB');
});

test('引擎包：在浏览器环境里求值后挂到 KimpoPlugins.attachment_docx_engine，版本 1.47.2', () => {
  const window = new Window({ url: 'https://kimpo.example.com/' });
  window.Vue = require('vue');
  window.process = undefined;
  const sandbox = window;
  vm.createContext(sandbox);
  vm.runInContext(engine, sandbox, { filename: 'engine.js' });
  const mod = sandbox.KimpoPlugins?.attachment_docx_engine;
  assert.equal(typeof mod?.createDocxEngine, 'function');
  assert.equal(mod.version, '1.47.2');
  assert.ok(window.document.querySelector('style[data-marker="kimpo-plugin-style-attachment_docx_engine"]'), 'SuperDoc 样式自注入');
  window.close();
});

test('引擎包：用宿主 Vue、遥测关闭、无开发分支残留、Chromium 88 垫片在位', () => {
  assert.doesNotMatch(engine, /__VUE_OPTIONS_API__\s*=|function createRenderer\(/, '不应打包第二份 Vue 运行时');
  assert.match(engine, /telemetry:\{enabled:!1\}/, '封装里遥测显式关闭');
  assert.doesNotMatch(engine, /process\$1\.env\.NODE_ENV/, 'SuperDoc 内联 process 垫片的开发分支应定为 production');
  for (const name of ['randomUUID', 'structuredClone', 'findLast', 'hasOwn']) assert.match(engine, new RegExp(name), `垫片 ${name}`);
});

test('随包许可与源码地址', () => {
  assert.match(readFileSync(join(dist, 'LICENSE'), 'utf8'), /GNU AFFERO GENERAL PUBLIC LICENSE/);
  assert.match(readFileSync(join(dist, 'superdoc', 'LICENSE'), 'utf8'), /GNU AFFERO GENERAL PUBLIC LICENSE/);
  assert.match(readFileSync(join(dist, 'SOURCE.txt'), 'utf8'), /https:\/\/github\.com\/somsne\/kimpo-attachment-docx/);
  assert.match(readFileSync(join(dist, 'superdoc', 'THIRD_PARTY.txt'), 'utf8'), /SuperDoc 1\.47\.2 — AGPL-3\.0/);
  assert.ok(!existsSync(join(dist, 'style.css')), '样式已内联，不应另出 style.css');
});
