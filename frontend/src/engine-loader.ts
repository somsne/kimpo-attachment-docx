/**
 * 编辑引擎的按需加载（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 进入编辑态时）。
 *
 * SuperDoc 体量大（数 MB），不进 plugin.umd.js：单独构建为随包文件 `superdoc/engine.js`（与 plugin.umd.js
 * 同级目录，见 vite.engine.config.ts），首次进入编辑态时插入一个 <script> 加载，加载后挂到
 * `window.KimpoPlugins.attachment_docx_engine`。基地址从插件脚本自身地址（document.currentScript.src，
 * 模块求值时取一次）推出，与 PDF 插件随包资源同一做法。查看态只用 docx-preview，不加载引擎。
 */
import type { DocxEngineModule } from './engine-types';

/** 随包引擎相对 plugin.umd.js 的路径。 */
export const ENGINE_PATH = 'superdoc/engine.js';
/** 引擎包挂载的全局名（window.KimpoPlugins 下）。 */
export const ENGINE_GLOBAL = 'attachment_docx_engine';

/**
 * 插件脚本自身地址：只能在脚本同步执行期间从 document.currentScript 取到，之后就是 null，
 * 所以在模块顶层取一次。
 */
const SCRIPT_SRC: string =
  typeof document !== 'undefined' && document.currentScript instanceof HTMLScriptElement ? document.currentScript.src : '';

/** 由插件脚本地址推出引擎地址；取不到（如单测里）得空串。只接受 http(s) 与 file（本地调试页）。 */
export function engineUrl(scriptSrc: string | null | undefined = SCRIPT_SRC): string {
  if (!scriptSrc) return '';
  try {
    const url = new URL(ENGINE_PATH, scriptSrc);
    return ['http:', 'https:', 'file:'].includes(url.protocol) ? url.href : '';
  } catch {
    return '';
  }
}

function readGlobal(): DocxEngineModule | undefined {
  const ns = (globalThis as { KimpoPlugins?: Record<string, unknown> }).KimpoPlugins;
  const mod = ns?.[ENGINE_GLOBAL] as DocxEngineModule | undefined;
  return mod && typeof mod.createDocxEngine === 'function' ? mod : undefined;
}

let loading: Promise<DocxEngineModule> | null = null;

/** 加载引擎（同一页面只加载一次；失败清缓存，下次可重试）。 */
export function loadDocxEngine(url: string = engineUrl()): Promise<DocxEngineModule> {
  const existing = readGlobal();
  if (existing) return Promise.resolve(existing);
  loading ??= new Promise<DocxEngineModule>((resolve, reject) => {
    if (!url) {
      reject(new Error('docx engine url unavailable'));
      return;
    }
    const script = document.createElement('script');
    script.src = url;
    script.async = true;
    script.dataset.kimpoPlugin = 'attachment_docx_engine';
    script.onload = () => {
      const mod = readGlobal();
      if (mod) resolve(mod);
      else reject(new Error('docx engine loaded but global missing'));
    };
    script.onerror = () => reject(new Error(`docx engine failed to load: ${url}`));
    document.head.appendChild(script);
  }).catch((error: unknown) => {
    loading = null;
    throw error;
  });
  return loading;
}

/** 单测用：清掉缓存。 */
export function resetDocxEngineLoader(): void {
  loading = null;
}
