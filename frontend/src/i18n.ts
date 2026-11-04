/**
 * 文档附件插件（kimpo-attachment-docx，填报态附件外壳内）的多语言薄封装。
 *
 * 查找链用 SDK createPluginI18n（插件语言包[当前] → 宿主中央包 → 插件语言包[默认] → key）；
 * 这里只包一层 Vue 响应式，让宿主切换语言时模板跟着刷新。宿主 i18n 能力取不到（老宿主 / 单测）时，
 * 按 ctx.env.locale 设定语言。
 */
import { ref } from 'vue';
import { createPluginI18n, type HostI18n } from '@kimposoft/plugin-sdk/i18n';
import zhCN from '../locales/zh-CN.json';
import enUS from '../locales/en-US.json';
import jaJP from '../locales/ja-JP.json';
import zhTW from '../locales/zh-TW.json';

/** 与 manifest.i18n.namespace 一致。 */
export const I18N_NAMESPACE = 'com.kimposoft.attachment-docx';

export type Translate = (key: string, params?: Record<string, unknown>) => string;

/** 从宿主桥取 i18n 能力；被裁剪的 Proxy 访问未声明键会抛错，一律视为没有。 */
export function readHostI18n(hostBridge: unknown): HostI18n | undefined {
  try {
    const candidate = (hostBridge as { i18n?: HostI18n } | null | undefined)?.i18n;
    return candidate && typeof candidate.t === 'function' && typeof candidate.onLocaleChange === 'function' ? candidate : undefined;
  } catch {
    return undefined;
  }
}

export function useAttachmentI18n(hostBridge: unknown, fallbackLocale?: string): { t: Translate; dispose: () => void } {
  const host = readHostI18n(hostBridge);
  const i18n = createPluginI18n({
    messages: { 'zh-CN': zhCN, 'en-US': enUS, 'ja-JP': jaJP, 'zh-TW': zhTW },
    defaultLocale: 'zh-CN',
    namespace: I18N_NAMESPACE,
    host
  });
  if (!host && fallbackLocale) i18n.setLocale(fallbackLocale);
  // 读一次 tick 建立依赖：语言切换时递增，模板里的 t() 随之重算。
  const tick = ref(0);
  const off = i18n.onChange(() => {
    tick.value++;
  });
  const t: Translate = (key, params) => {
    void tick.value;
    return i18n.t(key, params);
  };
  return {
    t,
    dispose: () => {
      off();
      i18n.dispose();
    }
  };
}
