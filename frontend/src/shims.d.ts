declare module '*.vue' {
  import type { DefineComponent } from 'vue';
  const component: DefineComponent<Record<string, unknown>, Record<string, unknown>, unknown>;
  export default component;
}

/** 插件版本，构建时由 vite define 注入（vite.config.ts）。 */
declare const __PLUGIN_VERSION__: string | undefined;
