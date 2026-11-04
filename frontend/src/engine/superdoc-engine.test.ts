// SuperDoc 引擎封装单测：mock superdoc 与 superdoc/ui，核对构造参数（遥测关闭、不挂自带工具栏）、导出不回落原文件、
// 「有无修改」按内容指纹判定（用户动手前的变化算载入期规范化）、上下标走 Document API、只读切 viewing。
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
  config: null as Record<string, any> | null,
  instance: null as any,
  editor: null as any,
  ui: null as any
}));

vi.mock('superdoc', () => ({
  SuperDoc: vi.fn(function (this: any, config: Record<string, any>) {
    state.config = config;
    this.setDocumentMode = vi.fn();
    this.destroy = vi.fn();
    this.export = vi.fn(async () => new Blob(['ORIGINAL']));
    state.instance = this;
  })
}));

vi.mock('superdoc/ui', () => ({
  createSuperDocUI: vi.fn(() => state.ui)
}));

import { createDocxEngine } from './superdoc-engine';

/** 可变的「文档」：eq 按 version 比较。 */
function makeDoc(version: number) {
  return { version, eq: (other: { version: number }) => other?.version === version, descendants: () => {} };
}

function setup() {
  const handlers: Record<string, Array<() => void>> = {};
  state.editor = {
    options: {},
    isEditable: true,
    state: { doc: makeDoc(1), selection: { empty: true, $from: { marks: () => [], depth: 0, node: () => ({ type: { name: 'doc' }, attrs: {} }) } } },
    on: vi.fn((event: string, fn: () => void) => {
      (handlers[event] ??= []).push(fn);
    }),
    off: vi.fn(),
    exportDocx: vi.fn(async () => new Blob(['EDITED'])),
    commands: { insertPageBreak: vi.fn(() => true), setStyleById: vi.fn(() => true) },
    doc: { format: { vertAlign: vi.fn() }, sections: { list: () => ({ items: [] }) } },
    converter: { linkedStyles: [] }
  };
  const observe = () => () => {};
  state.ui = {
    toolbar: { getSnapshot: () => ({ commands: { bold: { active: true, disabled: false } } }), observe, execute: vi.fn(() => true) },
    selection: { getSnapshot: () => ({ selectionTarget: { kind: 'selection' } }), observe },
    comments: { getSnapshot: () => ({ items: [] }), observe },
    trackChanges: { getSnapshot: () => ({ items: [] }), observe },
    zoom: { getSnapshot: () => ({ value: 100 }), observe },
    fonts: { getSnapshot: () => ({ options: [], sizeOptions: [] }), observe },
    viewport: { getScrollContainer: () => null },
    destroy: vi.fn()
  };
  const element = document.createElement('div');
  document.body.appendChild(element);
  const log = vi.fn();
  const engine = createDocxEngine({ element, file: new Blob(['x']), fileName: '合同.docx', user: { name: '张三' }, trackChanges: false, log });
  // 模拟 SuperDoc 回调：主体编辑器创建 → 就绪
  state.config!.onEditorCreate({ editor: state.editor });
  state.config!.onReady({ superdoc: state.instance });
  const emit = (event: string) => handlers[event]?.forEach(fn => fn());
  return { engine, element, log, emit };
}

beforeEach(() => {
  state.config = null;
});

afterEach(() => {
  document.body.innerHTML = '';
});

describe('createDocxEngine', () => {
  it('构造参数：遥测显式关闭、不挂 SuperDoc 自带工具栏、文档用 File 传入、署名用当前用户', async () => {
    const { engine } = setup();
    await engine.ready;
    expect(state.config!.telemetry).toEqual({ enabled: false });
    expect(state.config!.toolbar).toBeUndefined();
    expect(state.config!.modules.toolbar?.selector).toBeUndefined();
    expect(state.config!.document).toBeInstanceOf(File);
    expect((state.config!.document as File).name).toBe('合同.docx');
    expect(state.config!.user.name).toBe('张三');
    expect(state.config!.documentMode).toBe('editing');
    // 自带右键菜单是英文界面，关闭（插件自绘四语菜单）
    expect(state.config!.disableContextMenu).toBe(true);
  });

  it('导出直接调主编辑器 exportDocx，不用会回落原文件的 superdoc.export', async () => {
    const { engine } = setup();
    const blob = await engine.exportDocx();
    expect(await blob.text()).toBe('EDITED');
    expect(state.instance.export).not.toHaveBeenCalled();
    state.editor.exportDocx.mockResolvedValueOnce(undefined);
    await expect(engine.exportDocx()).rejects.toMatchObject({ code: 'EXPORT_EMPTY' });
  });

  it('有无修改按内容指纹：用户动手前的变化不算；动手后变了才算；保存后回到基线', async () => {
    const { engine, element, emit } = setup();
    const changes: boolean[] = [];
    engine.onContentChange(modified => changes.push(modified));
    // 载入期规范化：文档变了，但用户没动过
    state.editor.state.doc = makeDoc(2);
    emit('update');
    expect(changes).toEqual([]);
    // 用户按键后改了内容
    element.querySelector('.kad-superdoc-host')!.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true }));
    state.editor.state.doc = makeDoc(3);
    emit('update');
    expect(changes).toEqual([true]);
    // 撤销回到基线
    state.editor.state.doc = makeDoc(2);
    emit('update');
    expect(changes).toEqual([true, false]);
    // 再改、保存后重设基线
    state.editor.state.doc = makeDoc(4);
    emit('update');
    engine.markSaved();
    expect(changes).toEqual([true, false, true, false]);
  });

  it('上标：对当前选区调 Document API format.vertAlign；不在 SuperDoc 内置命令里的不转交工具栏', () => {
    const { engine } = setup();
    expect(engine.execute('superscript')).toBe(true);
    expect(state.editor.doc.format.vertAlign).toHaveBeenCalledWith({ target: { kind: 'selection' }, value: 'superscript' });
    expect(engine.execute('page-break')).toBe(true);
    expect(state.editor.commands.insertPageBreak).toHaveBeenCalled();
    engine.execute('bold');
    expect(state.ui.toolbar.execute).toHaveBeenCalledWith('bold');
    expect(engine.getSnapshot().commands.bold).toEqual({ active: true, disabled: false, value: undefined });
  });

  it('修订开关切 suggesting / editing；只读切 viewing 且期间修订开关不改模式', () => {
    const { engine } = setup();
    engine.setTrackChanges(true);
    expect(state.instance.setDocumentMode).toHaveBeenLastCalledWith('suggesting');
    engine.setReadOnly(true);
    expect(state.instance.setDocumentMode).toHaveBeenLastCalledWith('viewing');
    engine.setTrackChanges(false);
    expect(state.instance.setDocumentMode).toHaveBeenLastCalledWith('viewing');
    expect(engine.getSnapshot().trackChanges).toBe(false);
  });

  it('destroy 销毁 SuperDoc 与控制器并移除挂载容器', () => {
    const { engine, element } = setup();
    engine.destroy();
    expect(state.instance.destroy).toHaveBeenCalled();
    expect(state.ui.destroy).toHaveBeenCalled();
    expect(element.querySelector('.kad-superdoc-host')).toBeNull();
  });
});
