// DocxStage 组件测试：mock ctx（readSource 返回现拼的 docx、beginEdit 返回可控 session）、假编辑引擎（实现 DocxEngine）、
// mock docx-preview。覆盖查看、守门只读、编辑保存、往返自检拒存、导入自检转只读、失锁、修订接受 / 拒绝、批注回复。
import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from 'vitest';
import { flushPromises, mount, type VueWrapper } from '@vue/test-utils';
import type { AttachmentEditSession, AttachmentOpenContext, AttachmentOpenMode } from '@kimposoft/plugin-sdk/attachment-open';
import DocxStage from './DocxStage.vue';
import type { DocxEngine, DocxEngineModule, EngineSnapshot } from './engine-types';
import { summarizeDocumentXml, type DocxSummary } from './docx-summary';
import { documentXml, makeDocx, para } from './test-helpers';

vi.mock('docx-preview', () => ({
  renderAsync: vi.fn(async (_data: unknown, body: HTMLElement) => {
    body.innerHTML = '<div class="kad-docx-wrapper"><section class="kad-docx"><article><p>采购框架合同</p><p>正文</p></article></section><section class="kad-docx"><article><p>二、价格与付款</p></article></section></div>';
  })
}));

const DOC_XML = documentXml(para('采购框架合同', '<w:pStyle w:val="Heading1"/>') + para('正文') + para('二、价格与付款', '<w:pStyle w:val="Heading2"/>'));
const STYLES = '<w:styles><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style></w:styles>';

interface MockSession extends AttachmentEditSession {
  save: Mock;
  end: Mock;
}

function makeSession(): MockSession {
  return {
    id: 'sess_1',
    expiresAt: '2026-11-04T07:34:00Z',
    save: vi.fn(async (blob: Blob) => ({ assetId: 'ast_new', byteSize: blob.size })),
    end: vi.fn(async () => {})
  } as MockSession;
}

/** 假引擎：内容摘要与导出结果可控；onContentChange 由测试触发。 */
function makeEngine(options: { summary: DocxSummary; exportXml?: string }) {
  let contentCb: ((modified: boolean) => void) | null = null;
  const snapshot: EngineSnapshot = {
    ready: true,
    trackChanges: false,
    commands: {},
    styleId: null,
    comments: [{ id: 'c1', author: '张三', createdAt: '2026-11-04T06:52:00Z', text: '付款期限请确认。', resolved: false, parentId: null, quote: '30 个工作日' }],
    changes: [{ id: 'tc1', type: 'replace', author: '李四', createdAt: '2026-11-04T07:01:00Z', excerpt: '每年复核', deletedText: '不作调整', insertedText: '每年复核' }],
    activeCommentId: null,
    zoom: 100,
    pages: 2,
    currentPage: 1,
    words: 1286,
    fontFamilies: [],
    fontSizes: []
  };
  const engine = {
    ready: Promise.resolve(),
    getSnapshot: vi.fn(() => snapshot),
    subscribe: vi.fn(() => () => {}),
    onContentChange: vi.fn((cb: (modified: boolean) => void) => {
      contentCb = cb;
      return () => {
        contentCb = null;
      };
    }),
    markSaved: vi.fn(),
    execute: vi.fn(() => true),
    setTrackChanges: vi.fn(),
    setReadOnly: vi.fn(),
    styles: vi.fn(() => [{ id: 'Normal', name: 'Normal', builtin: true, headingLevel: 0 }]),
    applyStyle: vi.fn(() => true),
    outline: vi.fn(() => [{ level: 1, text: '采购框架合同', id: 'p1' }]),
    goToOutline: vi.fn(),
    goToPage: vi.fn(),
    find: vi.fn(() => ({ total: 0, current: 0 })),
    findNext: vi.fn(() => ({ total: 0, current: 0 })),
    findPrevious: vi.fn(() => ({ total: 0, current: 0 })),
    replaceCurrent: vi.fn(() => ({ total: 0, current: 0 })),
    replaceAll: vi.fn(() => 0),
    clearFind: vi.fn(),
    getPageSetup: vi.fn(() => null),
    setPageSetup: vi.fn(() => true),
    addComment: vi.fn(() => true),
    replyComment: vi.fn(() => true),
    resolveComment: vi.fn(() => true),
    deleteComment: vi.fn(() => true),
    focusComment: vi.fn(),
    acceptChange: vi.fn(() => true),
    rejectChange: vi.fn(() => true),
    acceptAll: vi.fn(() => true),
    rejectAll: vi.fn(() => true),
    focusChange: vi.fn(),
    exportDocx: vi.fn(async () => new Blob([(await makeDocx({ 'word/document.xml': options.exportXml ?? DOC_XML })) as BlobPart])),
    summarize: vi.fn(() => options.summary),
    destroy: vi.fn()
  } satisfies DocxEngine;
  const loadEngine = vi.fn(async (): Promise<DocxEngineModule> => ({ version: '1.47.2', createDocxEngine: () => engine }));
  return { engine, loadEngine, emitContent: (modified: boolean) => contentCb?.(modified) };
}

type MockCtx = { [K in keyof AttachmentOpenContext]: AttachmentOpenContext[K] extends (...args: never[]) => unknown ? Mock<AttachmentOpenContext[K]> : AttachmentOpenContext[K] };

async function makeCtx(options: { mode?: AttachmentOpenMode; canEdit?: boolean; canDownload?: boolean; documentXml?: string; extraParts?: Record<string, string | Uint8Array> } = {}) {
  const bytes = await makeDocx({ 'word/document.xml': options.documentXml ?? DOC_XML, 'word/styles.xml': STYLES, ...options.extraParts });
  let modeCb: ((mode: AttachmentOpenMode) => void) | null = null;
  let saveCb: (() => Promise<boolean>) | null = null;
  const session = makeSession();
  const ctx = {
    attachment: { assetId: 'ast_1', name: '采购框架合同-终稿.docx', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', byteSize: bytes.length },
    env: { locale: 'zh-CN', theme: 'light', timezone: 'Asia/Shanghai' },
    permissions: { canView: true, canDownload: options.canDownload ?? false, canEdit: options.canEdit ?? true },
    mode: options.mode ?? 'view',
    readSource: vi.fn(async () => ({ bytes: bytes.slice().buffer, contentType: '', truncated: false })),
    download: vi.fn(async () => {}),
    beginEdit: vi.fn(async () => session),
    setDirty: vi.fn(),
    setStatus: vi.fn(),
    onSourceChanged: vi.fn(() => () => {}),
    onModeChange: vi.fn((cb: (mode: AttachmentOpenMode) => void) => {
      modeCb = cb;
      return () => {
        modeCb = null;
      };
    }),
    onSaveRequest: vi.fn((cb: () => Promise<boolean>) => {
      saveCb = cb;
      return () => {
        if (saveCb === cb) saveCb = null;
      };
    }),
    log: vi.fn()
  };
  return {
    ctx: ctx as unknown as MockCtx,
    session,
    emitMode: (mode: AttachmentOpenMode) => modeCb?.(mode),
    requestSave: () => {
      if (!saveCb) throw new Error('插件未登记 onSaveRequest');
      return saveCb();
    }
  };
}

const mounted: VueWrapper[] = [];
function mountStage(ctx: MockCtx, loadEngine?: () => Promise<DocxEngineModule>) {
  const wrapper = mount(DocxStage, { props: { ctx: ctx as unknown as AttachmentOpenContext, loadEngine }, attachTo: document.body });
  mounted.push(wrapper);
  return wrapper;
}

/** JSZip 解压与 Blob 读取走宏任务，多轮让出直到界面稳定。 */
async function settle() {
  for (let i = 0; i < 25; i++) {
    await flushPromises();
    await new Promise(resolve => setTimeout(resolve, 0));
  }
}

beforeEach(() => {
  vi.useRealTimers();
});

afterEach(() => {
  mounted.splice(0).forEach(w => w.unmount());
});

const BASE = summarizeDocumentXml(DOC_XML);

describe('查看', () => {
  it('loading → 读全文 → ready，docx-preview 渲染两页，导航窗格列出标题', async () => {
    const h = await makeCtx();
    const wrapper = mountStage(h.ctx);
    expect(h.ctx.setStatus).toHaveBeenNthCalledWith(1, { state: 'loading' });
    await settle();
    expect(h.ctx.readSource).toHaveBeenCalledWith();
    expect(h.ctx.setStatus).toHaveBeenLastCalledWith({ state: 'ready' });
    // happy-dom 没有布局尺寸，当前页无从判断，只核总页数
    expect(wrapper.get('[data-testid="docx-page"]').text()).toMatch(/^第 \d \/ 2 页$/);
    const outline = wrapper.findAll('[data-testid="docx-outline"] button').map(b => b.text());
    expect(outline).toEqual(['采购框架合同', '二、价格与付款']);
    expect(wrapper.find('[data-testid="docx-editor"]').exists()).toBe(false);
    expect(h.ctx.beginEdit).not.toHaveBeenCalled();
    expect(h.ctx.log).toHaveBeenCalledWith('info', '文档附件已打开', expect.objectContaining({ headings: 2, guard: [] }));
  });

  it('下载按钮只在 canDownload 时出现，点击交给 ctx.download()', async () => {
    const hidden = mountStage((await makeCtx()).ctx);
    await settle();
    expect(hidden.find('[data-testid="docx-download"]').exists()).toBe(false);
    const h = await makeCtx({ canDownload: true });
    const shown = mountStage(h.ctx);
    await settle();
    await shown.get('[data-testid="docx-download"]').trigger('click');
    expect(h.ctx.download).toHaveBeenCalledTimes(1);
  });

  it('缩放按钮按档位放大缩小', async () => {
    const wrapper = mountStage((await makeCtx()).ctx);
    await settle();
    await wrapper.get('[data-testid="docx-zoom-in"]').trigger('click');
    expect(wrapper.get('[data-testid="docx-zoom"]').text()).toBe('110%');
    await wrapper.get('[data-testid="docx-zoom-out"]').trigger('click');
    await wrapper.get('[data-testid="docx-zoom-out"]').trigger('click');
    expect(wrapper.get('[data-testid="docx-zoom"]').text()).toBe('90%');
  });

  it('宿主拒绝 SOURCE_TOO_LARGE：单独提示文件过大', async () => {
    const h = await makeCtx();
    h.ctx.readSource.mockRejectedValueOnce(Object.assign(new Error('too large'), { code: 'SOURCE_TOO_LARGE' }));
    const wrapper = mountStage(h.ctx);
    await settle();
    expect(wrapper.get('[data-testid="docx-state"]').attributes('data-state')).toBe('too_large');
    expect(h.ctx.setStatus).toHaveBeenLastCalledWith({ state: 'error', message: '文件超过 50 MB，不能在线查看' });
  });

  it('「关于」给出版本、AGPL-3.0 许可与源码地址', async () => {
    const wrapper = mountStage((await makeCtx()).ctx);
    await settle();
    await wrapper.get('[data-testid="docx-about"]').trigger('click');
    const dialog = wrapper.get('[data-testid="docx-about-dialog"]');
    expect(dialog.text()).toContain('AGPL-3.0');
    expect(dialog.text()).toContain('SuperDoc 1.47.2');
    expect(dialog.get('[data-testid="docx-about-source"]').attributes('href')).toBe('https://github.com/somsne/kimpo-attachment-docx');
  });
});

describe('守门只读', () => {
  it('含文本框：出守门条幅，切到编辑也不取锁、不加载编辑器', async () => {
    const h = await makeCtx({ mode: 'edit', canDownload: true, documentXml: documentXml('<w:p><w:r><w:txbxContent>' + para('框') + '</w:txbxContent></w:r></w:p>') });
    const { loadEngine } = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, loadEngine);
    await settle();
    expect(wrapper.get('[data-testid="docx-guard"]').text()).toContain('此文件含文本框，在线编辑会丢失这些内容，已按只读打开');
    expect(h.ctx.beginEdit).not.toHaveBeenCalled();
    expect(loadEngine).not.toHaveBeenCalled();
    expect(h.ctx.log).toHaveBeenCalledWith('warn', expect.stringContaining('按只读打开'), expect.objectContaining({ reasons: ['textbox'] }));
    await wrapper.get('[data-testid="docx-guard-download"]').trigger('click');
    expect(h.ctx.download).toHaveBeenCalled();
  });

  it('复杂域：条幅列出域类型', async () => {
    const xml = documentXml('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF 1 = 1 "a" "b"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>');
    const wrapper = mountStage((await makeCtx({ documentXml: xml })).ctx);
    await settle();
    expect(wrapper.get('[data-testid="docx-guard"]').text()).toContain('复杂域代码（IF）');
  });
});

describe('编辑与保存', () => {
  it('取锁后才加载引擎；改动点亮未保存；宿主保存 → 往返自检通过 → session.save（文件名与类型沿用）', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    expect(h.ctx.beginEdit).toHaveBeenCalledTimes(1);
    expect(fake.loadEngine).toHaveBeenCalledTimes(1);
    expect(h.ctx.beginEdit.mock.invocationCallOrder[0]).toBeLessThan(fake.loadEngine.mock.invocationCallOrder[0]!);
    expect(wrapper.find('[data-testid="docx-ribbon"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="docx-status-page"]').text()).toBe('第 1 页 / 共 2 页');
    expect(wrapper.get('[data-testid="docx-status-words"]').text()).toBe('1,286 字');

    fake.emitContent(true);
    expect(h.ctx.setDirty).toHaveBeenLastCalledWith(true);

    await expect(h.requestSave()).resolves.toBe(true);
    expect(h.session.save).toHaveBeenCalledTimes(1);
    const [blob, opts] = h.session.save.mock.calls[0]!;
    expect(blob).toBeInstanceOf(Blob);
    expect(opts).toEqual({ contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', name: '采购框架合同-终稿.docx' });
    expect(h.ctx.setDirty).toHaveBeenLastCalledWith(false);
    expect(fake.engine.markSaved).toHaveBeenCalled();
    expect(h.ctx.log).toHaveBeenCalledWith('info', '附件已保存', expect.objectContaining({ assetId: 'ast_new' }));
  });

  it('没有改动时宿主请保存 → false，不调 session.save', async () => {
    const h = await makeCtx({ mode: 'edit' });
    mountStage(h.ctx, makeEngine({ summary: BASE }).loadEngine);
    await settle();
    await expect(h.requestSave()).resolves.toBe(false);
    expect(h.session.save).not.toHaveBeenCalled();
  });

  it('往返自检不通过：拒存（不调 session.save）、warn 日志、条幅，修改保留', async () => {
    const h = await makeCtx({ mode: 'edit', canDownload: true });
    // 导出结果丢了一段（模拟引擎导出丢内容）
    const fake = makeEngine({ summary: BASE, exportXml: documentXml(para('采购框架合同') + para('正文')) });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    fake.emitContent(true);
    await expect(h.requestSave()).rejects.toMatchObject({ code: 'ROUNDTRIP_CHECK_FAILED' });
    await settle();
    expect(h.session.save).not.toHaveBeenCalled();
    expect(h.ctx.log).toHaveBeenCalledWith('warn', '保存前往返自检未通过，已拒绝保存', expect.objectContaining({ paragraphs: [3, 2] }));
    expect(wrapper.find('[data-testid="docx-save-check"]').exists()).toBe(true);
    expect(h.ctx.setDirty).not.toHaveBeenLastCalledWith(false);
    expect(fake.engine.destroy).not.toHaveBeenCalled();
  });

  it('导入自检不通过（编辑器没装下原文件）：转只读、条幅，之后的改动不点亮未保存', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: summarizeDocumentXml(documentXml(para('采购框架合同'))) });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    expect(wrapper.find('[data-testid="docx-import-check"]').exists()).toBe(true);
    expect(fake.engine.setReadOnly).toHaveBeenCalledWith(true);
    expect(h.ctx.log).toHaveBeenCalledWith('warn', '编辑器导入自检未通过，转为只读', expect.any(Object));
    fake.emitContent(true);
    expect(h.ctx.setDirty).not.toHaveBeenCalledWith(true);
  });

  it('保存时失锁：保留编辑器与内容、出失锁条幅；「重新获取并保存」重新取锁后写出', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    h.session.save.mockRejectedValueOnce(Object.assign(new Error('lost'), { code: 'EDIT_LOCK_LOST' }));
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    fake.emitContent(true);
    await expect(h.requestSave()).rejects.toMatchObject({ code: 'EDIT_LOCK_LOST' });
    await settle();
    expect(wrapper.find('[data-testid="docx-lock-lost"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="docx-editor"]').exists()).toBe(true);
    expect(fake.engine.destroy).not.toHaveBeenCalled();
    await wrapper.get('[data-testid="docx-relock"]').trigger('click');
    await settle();
    expect(h.ctx.beginEdit).toHaveBeenCalledTimes(2);
    expect(h.session.save).toHaveBeenCalledTimes(2);
    expect(wrapper.find('[data-testid="docx-lock-lost"]').exists()).toBe(false);
  });

  it('编辑锁被占用：没有确认框能力时出占用条幅，不加载编辑器', async () => {
    const h = await makeCtx({ mode: 'edit' });
    h.ctx.beginEdit.mockRejectedValueOnce(Object.assign(new Error('locked'), { code: 'EDIT_LOCKED', data: { lockedBy: { displayName: '李四', since: '2026-11-04T07:02:00Z' } } }));
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    expect(wrapper.get('[data-testid="docx-lock-held"]').text()).toContain('正由 李四 编辑');
    expect(fake.loadEngine).not.toHaveBeenCalled();
  });

  it('切回查看：结束会话、卸载编辑器并销毁引擎', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    h.emitMode('view');
    await settle();
    expect(h.session.end).toHaveBeenCalledWith({ discard: false });
    expect(fake.engine.destroy).toHaveBeenCalled();
    expect(wrapper.find('[data-testid="docx-viewer"]').exists()).toBe(true);
  });
});

describe('审阅栏', () => {
  it('修订卡片显示「删除…，插入…」，接受 / 拒绝转交引擎；功能区「全部接受」', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    const card = wrapper.get('[data-testid="docx-change"]');
    expect(card.text()).toContain('李四 · 修订');
    expect(card.text()).toContain('删除「不作调整」，插入「每年复核」');
    expect(card.text()).toContain('15:01');
    await card.get('[data-testid="docx-change-accept"]').trigger('click');
    expect(fake.engine.acceptChange).toHaveBeenCalledWith('tc1');
    await card.get('[data-testid="docx-change-reject"]').trigger('click');
    expect(fake.engine.rejectChange).toHaveBeenCalledWith('tc1');
    await wrapper.get('[data-testid="docx-tab-review"]').trigger('click');
    await wrapper.get('[data-testid="docx-accept-all"]').trigger('click');
    expect(fake.engine.acceptAll).toHaveBeenCalled();
  });

  it('批注：空输入时按钮是「解决」，输入后变「回复」并转交引擎', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    const card = wrapper.get('[data-testid="docx-comment"]');
    expect(card.get('[data-testid="docx-comment-action"]').text()).toBe('解决');
    await card.get('input').setValue('已与财务确认');
    expect(card.get('[data-testid="docx-comment-action"]').text()).toBe('回复');
    await card.get('[data-testid="docx-comment-action"]').trigger('click');
    expect(fake.engine.replyComment).toHaveBeenCalledWith('c1', '已与财务确认');
    await card.get('[data-testid="docx-comment-action"]').trigger('click');
    expect(fake.engine.resolveComment).toHaveBeenCalledWith('c1');
  });

  it('修订开关转交引擎', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    await wrapper.get('[data-testid="docx-track-toggle"] input').setValue(true);
    expect(fake.engine.setTrackChanges).toHaveBeenCalledWith(true);
  });

  it('功能区按钮转交引擎命令（加粗、上标、两端对齐、样式库）', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    await wrapper.get('[data-testid="docx-cmd-bold"]').trigger('click');
    await wrapper.get('[data-testid="docx-cmd-superscript"]').trigger('click');
    await wrapper.get('[data-testid="docx-align-justify"]').trigger('click');
    expect(fake.engine.execute).toHaveBeenCalledWith('bold', undefined);
    expect(fake.engine.execute).toHaveBeenCalledWith('superscript', undefined);
    expect(fake.engine.execute).toHaveBeenCalledWith('text-align', 'justify');
    await wrapper.get('[data-style-id="Normal"]').trigger('click');
    expect(fake.engine.applyStyle).toHaveBeenCalledWith('Normal');
    expect(wrapper.get('[data-style-id="Normal"]').text()).toContain('正文');
  });

  it('右键菜单（插件自绘、四语）：通用项 + 表格内才出行列操作，点「新建批注」打开批注输入框', async () => {
    const h = await makeCtx({ mode: 'edit' });
    const fake = makeEngine({ summary: BASE });
    const wrapper = mountStage(h.ctx, fake.loadEngine);
    await settle();
    const pages = wrapper.get('.kad-pages');
    await pages.trigger('contextmenu', { clientX: 100, clientY: 100 });
    const keys = () => wrapper.findAll('[data-testid="docx-context-menu"] [data-menu-key]').map(b => b.attributes('data-menu-key'));
    expect(keys()).toEqual(['cut', 'copy', 'paste', 'comment']);
    expect(wrapper.get('[data-menu-key="cut"]').text()).toBe('剪切');
    await wrapper.get('[data-menu-key="comment"]').trigger('click');
    expect(wrapper.find('[data-testid="docx-context-menu"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="docx-new-comment-box"]').exists()).toBe(true);
    // 光标在表格里、选区有修订
    fake.engine.getSnapshot().commands['table-delete-row'] = { active: false, disabled: false };
    fake.engine.getSnapshot().commands['track-changes-accept-selection'] = { active: false, disabled: false };
    wrapper.findComponent({ name: 'DocxEditor' }).vm.snapshot = { ...fake.engine.getSnapshot() };
    await pages.trigger('contextmenu', { clientX: 100, clientY: 100 });
    expect(keys()).toContain('table-add-row-after');
    expect(keys()).toContain('accept');
    await wrapper.get('[data-menu-key="table-add-row-after"]').trigger('click');
    expect(fake.engine.execute).toHaveBeenCalledWith('table-add-row-after');
  });
});
