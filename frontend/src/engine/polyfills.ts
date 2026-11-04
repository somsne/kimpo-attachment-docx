/**
 * 编辑引擎的最小运行时垫片（文档附件插件 kimpo-attachment-docx，只随 engine.js 加载）。
 *
 * 浏览器下限 Chromium 88（信创浏览器内核）；SuperDoc 1.47.2 的构建产物用到少量 88 之后才有的内置方法。
 * 只在缺失时补，且只补构建产物里实际用到的（scripts/bundle.test.mjs 核对清单）；查看态不加载本文件。
 */
function define(target: object, name: string, value: unknown): void {
  if (!(name in target)) Object.defineProperty(target, name, { value, writable: true, configurable: true, enumerable: false });
}

function at(this: { length: number; [index: number]: unknown }, index: number): unknown {
  const n = Math.trunc(index) || 0;
  const i = n < 0 ? this.length + n : n;
  return i < 0 || i >= this.length ? undefined : this[i];
}

define(Array.prototype, 'at', at);
define(String.prototype, 'at', function (this: string, index: number) {
  return at.call(this as unknown as { length: number }, index);
});
const typedArray = Object.getPrototypeOf(Int8Array.prototype) as object;
define(typedArray, 'at', at);

define(Array.prototype, 'findLast', function <T>(this: T[], fn: (v: T, i: number, a: T[]) => unknown, thisArg?: unknown) {
  for (let i = this.length - 1; i >= 0; i--) if (fn.call(thisArg, this[i]!, i, this)) return this[i];
  return undefined;
});
define(Array.prototype, 'findLastIndex', function <T>(this: T[], fn: (v: T, i: number, a: T[]) => unknown, thisArg?: unknown) {
  for (let i = this.length - 1; i >= 0; i--) if (fn.call(thisArg, this[i]!, i, this)) return i;
  return -1;
});
define(Object, 'hasOwn', (obj: object, key: PropertyKey) => Object.prototype.hasOwnProperty.call(obj, key));

/**
 * crypto.randomUUID：Chromium 92 起才有，且只在安全上下文（https / localhost）里提供——内网 http 部署的 Kimpo
 * 即使是新浏览器也取不到，SuperDoc 生成段落 / 批注 ID 要用它。
 */
const cryptoObj = (globalThis as { crypto?: Crypto }).crypto;
if (cryptoObj && typeof cryptoObj.getRandomValues === 'function' && typeof (cryptoObj as { randomUUID?: unknown }).randomUUID !== 'function') {
  define(cryptoObj, 'randomUUID', () => {
    const b = cryptoObj.getRandomValues(new Uint8Array(16));
    b[6] = (b[6]! & 0x0f) | 0x40;
    b[8] = (b[8]! & 0x3f) | 0x80;
    const h = Array.from(b, x => x.toString(16).padStart(2, '0')).join('');
    return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
  });
}

/** structuredClone：Chromium 98 起才有。补一个覆盖 SuperDoc 用法（普通对象 / 数组 / Date / Map / Set / 二进制，含循环引用）的深拷贝。 */
export function structuredClonePolyfill<T>(value: T): T {
  const seen = new Map<unknown, unknown>();
  const clone = (input: unknown): unknown => {
    if (input === null || typeof input !== 'object') {
      if (typeof input === 'function' || typeof input === 'symbol') throw new DOMException('value could not be cloned', 'DataCloneError');
      return input;
    }
    if (seen.has(input)) return seen.get(input);
    if (input instanceof Date) return new Date(input.getTime());
    if (input instanceof RegExp) return new RegExp(input.source, input.flags);
    if (input instanceof ArrayBuffer) return input.slice(0);
    if (ArrayBuffer.isView(input)) {
      const view = input as unknown as { constructor: new (b: ArrayBuffer, o: number, l: number) => unknown; buffer: ArrayBuffer; byteOffset: number; byteLength: number; length?: number };
      const copy = view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength);
      return input instanceof DataView ? new DataView(copy) : new view.constructor(copy, 0, view.length ?? view.byteLength);
    }
    if (input instanceof Map) {
      const out = new Map();
      seen.set(input, out);
      input.forEach((v, k) => out.set(clone(k), clone(v)));
      return out;
    }
    if (input instanceof Set) {
      const out = new Set();
      seen.set(input, out);
      input.forEach(v => out.add(clone(v)));
      return out;
    }
    if (typeof Blob !== 'undefined' && input instanceof Blob) return input;
    const out: Record<string, unknown> | unknown[] = Array.isArray(input) ? [] : {};
    seen.set(input, out);
    for (const key of Object.keys(input)) (out as Record<string, unknown>)[key] = clone((input as Record<string, unknown>)[key]);
    return out;
  };
  return clone(value) as T;
}
define(globalThis, 'structuredClone', structuredClonePolyfill);
