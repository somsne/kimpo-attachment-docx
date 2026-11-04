/**
 * 两份构建配置（vite.config.ts 主包、vite.engine.config.ts 编辑引擎包）共用的 vite 插件。
 */
import type { Plugin } from 'vite';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

/** 源码仓库地址（AGPL-3.0 源码提供义务：随包文件、README、插件「关于」都给这个地址）。 */
export const SOURCE_REPO = 'https://github.com/somsne/kimpo-attachment-docx';

/**
 * cssInjectInline —— 把构建产出的 CSS 内联进入口 chunk，运行时自注入 <style>。
 * 宿主 loader 只用 <script> 加载 plugin.umd.js、不加载伴随的 style.css（与登录页插件同口径）；
 * 引擎包同理（engine-loader.ts 只插 <script>）。
 */
export function cssInjectInline(marker: string): Plugin {
  return {
    name: 'kimpo-css-inject-inline',
    apply: 'build',
    enforce: 'post',
    generateBundle(_options, bundle) {
      let css = '';
      for (const [fileName, asset] of Object.entries(bundle)) {
        if (asset.type === 'asset' && fileName.endsWith('.css')) {
          css += typeof asset.source === 'string' ? asset.source : Buffer.from(asset.source).toString('utf8');
          delete bundle[fileName];
        }
      }
      if (!css) return;
      const id = `kimpo-plugin-style-${marker}`;
      const snippet = `\n(function(){try{if(typeof document==="undefined")return;`
        + `if(document.querySelector("style[data-marker=\\"${id}\\"]"))return;`
        + `var s=document.createElement("style");s.setAttribute("data-marker","${id}");`
        + `s.textContent=${JSON.stringify(css)};document.head.appendChild(s);}catch(e){}})();`;
      for (const chunk of Object.values(bundle)) {
        if (chunk.type === 'chunk' && chunk.isEntry) chunk.code += snippet;
      }
    }
  };
}

/** 随包许可文件：把若干源文件按目标路径写进产物目录。 */
export function emitLicenseFiles(files: Record<string, string | (() => string)>): Plugin {
  return {
    name: 'kimpo-emit-license-files',
    apply: 'build',
    generateBundle() {
      for (const [fileName, source] of Object.entries(files)) {
        this.emitFile({ type: 'asset', fileName, source: typeof source === 'function' ? source() : readFileSync(source) });
      }
    }
  };
}

/** 本仓 LICENSE（AGPL-3.0 全文）路径。 */
export const REPO_LICENSE = resolve(__dirname, '..', 'LICENSE');
