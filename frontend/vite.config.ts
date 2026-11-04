import { defineConfig } from 'vitest/config';
import vue from '@vitejs/plugin-vue';
import { resolve } from 'node:path';
import { cssInjectInline, emitLicenseFiles, REPO_LICENSE, SOURCE_REPO } from './vite.shared';
import pkg from './package.json';

/**
 * 文档附件插件（kimpo-attachment-docx）前端主包 UMD 构建：查看态（docx-preview）、守门、编辑态界面。
 * 外部化 vue（宿主 window.Vue）；@kimposoft/plugin-sdk、docx-preview、JSZip 打进 bundle。
 * SuperDoc 编辑引擎不在本包：单独构建为 dist/superdoc/engine.js（vite.engine.config.ts），进入编辑态才加载。
 * 产物 frontend/dist/plugin.umd.js + dist/LICENSE + dist/SOURCE.txt；manifest 声明包内 canonical 路径
 * frontend/plugin.umd.js，打包时整个 dist/ 目录内容收进 KPP 的 frontend/（保持同级结构）。
 */
export default defineConfig({
  plugins: [
    vue(),
    cssInjectInline('attachment_docx'),
    emitLicenseFiles({
      LICENSE: REPO_LICENSE,
      'SOURCE.txt': () =>
        `kimpo-attachment-docx ${pkg.version} (AGPL-3.0-only)\n\n`
        + `Complete corresponding source code: ${SOURCE_REPO}\n`
        + `You may obtain the source of this exact version from the tag or commit matching the plugin version.\n\n`
        + `完整对应源码：${SOURCE_REPO}\n本插件以 AGPL-3.0 发布，向任何获得本插件的人提供源码；编辑引擎 superdoc/engine.js 的许可见 superdoc/LICENSE。\n`
    })
  ],
  define: { __PLUGIN_VERSION__: JSON.stringify(pkg.version) },
  build: {
    target: 'chrome88',
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'KimpoPlugins.attachment_docx',
      fileName: () => 'plugin.umd.js',
      formats: ['umd']
    },
    rollupOptions: {
      external: ['vue'],
      output: { globals: { vue: 'Vue' } }
    }
  },
  test: {
    environment: 'happy-dom',
    include: ['src/**/*.test.ts'],
    testTimeout: 60000
  }
});
