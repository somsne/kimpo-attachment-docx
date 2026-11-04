// 杂项纯函数：字数统计、字号名、样式显示名、四语文案齐全与占位参数一致、引擎地址推导与加载。
import { afterEach, describe, expect, it } from 'vitest';
import { countWords } from './word-count';
import { fontSizeLabel, fontSizeOptions, styleDisplayName, styleSubtitle } from './ribbon-data';
import { engineUrl, loadDocxEngine, resetDocxEngineLoader } from './engine-loader';
import zhCN from '../locales/zh-CN.json';
import zhTW from '../locales/zh-TW.json';
import enUS from '../locales/en-US.json';
import jaJP from '../locales/ja-JP.json';

const t = (key: string, params?: Record<string, unknown>) => `${key}${params ? JSON.stringify(params) : ''}`;

describe('countWords', () => {
  it('中文逐字计、英文按词计、标点不计', () => {
    expect(countWords('采购框架合同，第 2 条。')).toBe(9);
    expect(countWords("Hello world, it's 2026.")).toBe(4);
    expect(countWords('')).toBe(0);
  });
});

describe('ribbon-data', () => {
  it('中文界面字号显示中文字号名，其他语言显示数字', () => {
    expect(fontSizeLabel('12pt', 'zh-CN')).toBe('小四');
    expect(fontSizeLabel('14pt', 'zh-TW')).toBe('四號');
    expect(fontSizeLabel('12pt', 'en-US')).toBe('12');
    expect(fontSizeLabel('13pt', 'zh-CN')).toBe('13');
    expect(fontSizeLabel('', 'zh-CN')).toBe('');
    expect(fontSizeOptions('zh-CN')[0]).toEqual({ label: '初号', value: '42pt' });
    expect(fontSizeOptions('en-US')[0]?.label).toBe('8');
  });

  it('样式库：预置样式翻译、自定义样式用原名并标「自定义」', () => {
    expect(styleDisplayName({ name: 'heading 2' }, t)).toBe('docx.style.heading{"n":"2"}');
    expect(styleDisplayName({ name: 'Normal' }, t)).toBe('docx.style.normal');
    expect(styleDisplayName({ name: '条款正文' }, t)).toBe('条款正文');
    expect(styleSubtitle({ id: 'c', name: '条款正文', builtin: false, headingLevel: 0 }, t, 'zh-CN')).toBe('docx.style.custom');
    expect(styleSubtitle({ id: 'h', name: 'heading 1', builtin: true, headingLevel: 1, fontFamily: '"黑体", sans-serif', fontSizePt: 22 }, t, 'zh-CN')).toBe('黑体 二号');
  });
});

describe('locales', () => {
  const all = { 'zh-TW': zhTW, 'en-US': enUS, 'ja-JP': jaJP } as Record<string, Record<string, string>>;
  const placeholders = (s: string) => (s.match(/\{\w+\}/g) ?? []).sort().join(',');

  it.each(Object.keys(all))('%s 与 zh-CN 键集合相同、占位参数一致、无空串', name => {
    const other = all[name]!;
    expect(Object.keys(other).sort()).toEqual(Object.keys(zhCN).sort());
    for (const [key, value] of Object.entries(zhCN as Record<string, string>)) {
      expect(other[key], key).toBeTruthy();
      expect(placeholders(other[key]!), key).toBe(placeholders(value));
    }
  });
});

describe('engine-loader', () => {
  afterEach(() => {
    resetDocxEngineLoader();
    delete (globalThis as { KimpoPlugins?: unknown }).KimpoPlugins;
    document.head.querySelectorAll('script[data-kimpo-plugin]').forEach(s => s.remove());
  });

  it('从插件脚本地址推出同级 superdoc/engine.js；非 http(s) / file 协议不给地址', () => {
    expect(engineUrl('https://kimpo.example.com/plugins/x/frontend/plugin.umd.js?v=1')).toBe('https://kimpo.example.com/plugins/x/frontend/superdoc/engine.js');
    expect(engineUrl('data:text/javascript,1')).toBe('');
    expect(engineUrl('')).toBe('');
  });

  it('已加载过（全局在位）直接返回，不再插脚本', async () => {
    const mod = { version: '1.47.2', createDocxEngine: () => ({}) };
    (globalThis as { KimpoPlugins?: unknown }).KimpoPlugins = { attachment_docx_engine: mod };
    await expect(loadDocxEngine('https://x/engine.js')).resolves.toBe(mod);
    expect(document.head.querySelector('script[data-kimpo-plugin]')).toBeNull();
  });

  it('取不到地址时拒绝，且失败后可重试', async () => {
    await expect(loadDocxEngine('')).rejects.toThrow(/url unavailable/);
    const mod = { version: '1.47.2', createDocxEngine: () => ({}) };
    (globalThis as { KimpoPlugins?: unknown }).KimpoPlugins = { attachment_docx_engine: mod };
    await expect(loadDocxEngine('')).resolves.toBe(mod);
  });
});
