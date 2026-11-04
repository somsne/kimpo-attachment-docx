import { defineConfig, type Plugin } from 'vite';
import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { cssInjectInline, emitLicenseFiles, SOURCE_REPO } from './vite.shared';
import pkg from './package.json';

/** superdoc 的 exports 不含 package.json，按 node_modules 路径直接取。 */
const superdocRoot = resolve(__dirname, 'node_modules/superdoc');
const superdocPkg = JSON.parse(readFileSync(join(superdocRoot, 'package.json'), 'utf8')) as { version: string };

/**
 * SuperDoc 1.47.2 的发布产物里内联了 npm `process` 浏览器垫片（env 为空对象），其中 91 处
 * `process$1.env.NODE_ENV` 判断（内联的 Pinia、Vue 工具代码）因此全走开发分支：Pinia 会往页面全局
 * __VUE_DEVTOOLS_PLUGINS__ 塞对象，再被 SuperDoc 自己的 disablePiniaDevtools 清理逻辑按数组解构而崩溃。
 * 构建时把这个判断定为 "production"（只动 superdoc/dist 下的文件；bundle 测试核对已无残留）。
 */
function superdocProductionEnv(): Plugin {
  return {
    name: 'kimpo-superdoc-production-env',
    enforce: 'pre',
    transform(code, id) {
      if (!/[\\/]superdoc[\\/]dist[\\/]/.test(id) || !code.includes('process$1.env.NODE_ENV')) return null;
      return { code: code.split('process$1.env.NODE_ENV').join('"production"'), map: null };
    }
  };
}

/**
 * 文档附件插件的编辑引擎包（随包文件 dist/superdoc/engine.js）：SuperDoc 1.47.2（AGPL-3.0）+ 引擎封装。
 * IIFE，挂到 window.KimpoPlugins.attachment_docx_engine；vue 外部化（宿主 window.Vue），其余依赖与 SuperDoc 样式
 * 打进本文件（样式运行时自注入）。由 plugin.umd.js 在进入编辑态时按需加载（engine-loader.ts）。
 * 同目录写出 SuperDoc 的 LICENSE 与 THIRD_PARTY.txt（依赖许可清单与源码地址）。
 * 目标 chrome88（语法降级）；88 之后才有的内置方法由 src/engine/polyfills.ts 按需补。
 */
export default defineConfig({
  plugins: [
    superdocProductionEnv(),
    cssInjectInline('attachment_docx_engine'),
    emitLicenseFiles({
      'superdoc/LICENSE': join(superdocRoot, 'LICENSE'),
      'superdoc/THIRD_PARTY.txt': () =>
        `kimpo-attachment-docx ${pkg.version} — editing engine bundle (engine.js)\n\n`
        + `SuperDoc ${superdocPkg.version} — AGPL-3.0 — https://github.com/superdoc/docx-editor (license text: ./LICENSE)\n`
        + `Code inlined in SuperDoc's own dist (shipped under SuperDoc's license notice): ProseMirror (MIT), Pinia (MIT), `
        + `Konva (MIT), JSZip (MIT OR GPL-3.0-or-later, used under MIT), xml-js (MIT), uuid (MIT), eventemitter3 (MIT), process / buffer shims (MIT).\n`
        + `Imported by SuperDoc and bundled here: yjs 13.6.33 (MIT), y-protocols 1.0.7 (MIT), lib0 0.2.119 (MIT), `
        + `@hocuspocus/provider 2.15.3 and @hocuspocus/common 2.15.3 (MIT; collaboration code, not used by this plugin).\n`
        + `Engine wrapper source (AGPL-3.0): ${SOURCE_REPO}\n`
    })
  ],
  define: {
    'process.env.NODE_ENV': JSON.stringify('production'),
    __VUE_PROD_DEVTOOLS__: 'false'
  },
  build: {
    target: 'chrome88',
    // 排查引擎问题时 KAD_ENGINE_DEBUG=1 出未压缩版
    minify: process.env.KAD_ENGINE_DEBUG ? false : 'esbuild',
    outDir: 'dist',
    emptyOutDir: false,
    // 引擎包按需加载，体量大属预期
    chunkSizeWarningLimit: 16384,
    lib: {
      entry: resolve(__dirname, 'src/engine/index.ts'),
      name: 'KimpoPlugins.attachment_docx_engine',
      fileName: () => 'superdoc/engine.js',
      formats: ['iife']
    },
    rollupOptions: {
      external: ['vue'],
      output: { globals: { vue: 'Vue' }, inlineDynamicImports: true }
    }
  }
});
