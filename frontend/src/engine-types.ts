/**
 * 编辑引擎接口（文档附件插件 kimpo-attachment-docx，填报态附件外壳内 · 编辑态）。
 *
 * 编辑态界面（DocxEditor.vue，自绘功能区 / 导航 / 审阅栏 / 状态栏）只经本接口驱动引擎；真正的实现在
 * src/engine/（SuperDoc 1.47.2，AGPL-3.0），单独构建为随包文件 dist/superdoc/engine.js，进入编辑态才加载
 * （见 engine-loader.ts）。单测用假引擎实现本接口。
 */
import type { DocxSummary } from './docx-summary';

/** 功能区命令（与 SuperDoc 内置命令 ID 同名的直接转交；其余由引擎自己实现）。 */
export type DocxCommandId =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'superscript'
  | 'subscript'
  | 'font-family'
  | 'font-size'
  | 'font-grow'
  | 'font-shrink'
  | 'text-color'
  | 'highlight-color'
  | 'clear-formatting'
  | 'copy-format'
  | 'bullet-list'
  | 'numbered-list'
  | 'multilevel-list'
  | 'indent-increase'
  | 'indent-decrease'
  | 'text-align'
  | 'line-height'
  | 'formatting-marks'
  | 'undo'
  | 'redo'
  | 'cut'
  | 'copy'
  | 'paste'
  | 'table-insert'
  | 'image'
  | 'page-break'
  | 'section-break'
  | 'table-of-contents-insert'
  | 'link'
  | 'table-add-row-before'
  | 'table-add-row-after'
  | 'table-delete-row'
  | 'table-add-column-before'
  | 'table-add-column-after'
  | 'table-delete-column'
  | 'table-delete'
  | 'table-merge-cells'
  | 'table-split-cell'
  | 'track-changes-accept-selection'
  | 'track-changes-reject-selection'
  | 'ruler'
  | 'zoom'
  | 'zoom-fit-width';

export interface CommandState {
  active: boolean;
  disabled: boolean;
  value?: unknown;
}

/** 样式库条目（取文档自带样式，含自定义样式）。 */
export interface StyleOption {
  id: string;
  name: string;
  /** 是否内置（Word 预置样式名）；自定义样式为 false。 */
  builtin: boolean;
  /** 标题级别（0 = 非标题）。 */
  headingLevel: number;
  /** 预览用的字号（pt）、字体、加粗，缺省时由界面按普通正文展示。 */
  fontSizePt?: number;
  fontFamily?: string;
  bold?: boolean;
}

export interface EngineOutlineItem {
  level: number;
  text: string;
  /** 引擎内定位用的标识（段落 ID）。 */
  id: string;
}

export interface EngineComment {
  id: string;
  author: string;
  /** UTC ISO；显示时按用户时区折算。 */
  createdAt: string | null;
  text: string;
  resolved: boolean;
  parentId: string | null;
  /** 批注锚定的原文片段（可能为空）。 */
  quote: string;
}

export interface EngineChange {
  id: string;
  /** replace = 删除并插入（Word 的替换修订对）。 */
  type: 'insert' | 'delete' | 'replace' | 'format' | 'other';
  author: string;
  createdAt: string | null;
  /** 修订涉及的文字片段。 */
  excerpt: string;
  /** 替换修订的删除部分与插入部分。 */
  deletedText?: string;
  insertedText?: string;
}

export interface PageSetup {
  orientation: 'portrait' | 'landscape';
  /** 纸张宽高（英寸）。 */
  width: number;
  height: number;
  /** 页边距（英寸）。 */
  margins: { top: number; right: number; bottom: number; left: number };
}

export interface FindResult {
  total: number;
  /** 当前命中序号（从 1 起；没有命中为 0）。 */
  current: number;
}

/** 引擎状态快照（界面订阅后按需重绘）。 */
export interface EngineSnapshot {
  ready: boolean;
  /** 修订（建议）模式是否开着。 */
  trackChanges: boolean;
  commands: Partial<Record<DocxCommandId, CommandState>>;
  /** 当前段落样式 ID。 */
  styleId: string | null;
  comments: EngineComment[];
  changes: EngineChange[];
  activeCommentId: string | null;
  zoom: number;
  pages: number;
  currentPage: number;
  words: number;
  /** 引擎声明可用的字体与字号选项。 */
  fontFamilies: Array<{ label: string; value: string }>;
  fontSizes: Array<{ label: string; value: string }>;
}

export interface DocxEngineOptions {
  /** 挂载容器（引擎在里面画页面）。 */
  element: HTMLElement;
  /** 原文件。 */
  file: Blob;
  fileName: string;
  /** 当前用户（批注与修订署名）。 */
  user: { name: string; email?: string };
  /** 初始修订开关。 */
  trackChanges: boolean;
  /** 关键位置日志（转交 ctx.log）。 */
  log: (level: 'info' | 'warn' | 'error', message: string, data?: Record<string, unknown>) => void;
}

export interface DocxEngine {
  /** 文档就绪（已导入、已排版）后 resolve；导入失败 reject。 */
  ready: Promise<void>;
  getSnapshot(): EngineSnapshot;
  /** 状态变化（选区、命令状态、批注修订、页数等）时回调；返回取消订阅。 */
  subscribe(listener: (snapshot: EngineSnapshot) => void): () => void;
  /**
   * 「有无未保存修改」变化时回调（用于 setDirty）：引擎按内容指纹（正文 + 批注）与基线比对，
   * 载入期的规范化事务不算修改；返回取消订阅。
   */
  onContentChange(listener: (modified: boolean) => void): () => void;
  /** 保存成功后把当前内容记为新基线。 */
  markSaved(): void;
  execute(id: DocxCommandId, payload?: unknown): boolean;
  setTrackChanges(on: boolean): void;
  /** 只读（往返自检未过等）：切到 SuperDoc 的 viewing 模式，不再接受修改。 */
  setReadOnly(on: boolean): void;
  styles(): StyleOption[];
  applyStyle(styleId: string): boolean;
  outline(): EngineOutlineItem[];
  goToOutline(id: string): void;
  goToPage(page: number): void;
  find(query: string, options?: { caseSensitive?: boolean }): FindResult;
  findNext(): FindResult;
  findPrevious(): FindResult;
  replaceCurrent(replacement: string): FindResult;
  replaceAll(query: string, replacement: string): number;
  clearFind(): void;
  getPageSetup(): PageSetup | null;
  setPageSetup(patch: Partial<Pick<PageSetup, 'orientation'>> & { paper?: 'A4' | 'A3' | 'Letter' | 'Legal' | 'B5'; margins?: PageSetup['margins'] }): boolean;
  addComment(text: string): boolean;
  replyComment(parentId: string, text: string): boolean;
  resolveComment(id: string): boolean;
  deleteComment(id: string): boolean;
  focusComment(id: string): void;
  acceptChange(id: string): boolean;
  rejectChange(id: string): boolean;
  acceptAll(): boolean;
  rejectAll(): boolean;
  focusChange(id: string): void;
  /** 导出当前内容为 docx（失败抛错，不回落原文件）。 */
  exportDocx(): Promise<Blob>;
  /** 编辑器模型的往返摘要（口径同 docx-summary）。 */
  summarize(): DocxSummary;
  destroy(): void;
}

/** 引擎包导出（dist/superdoc/engine.js 挂在 window.KimpoPlugins.attachment_docx_engine）。 */
export interface DocxEngineModule {
  version: string;
  createDocxEngine(options: DocxEngineOptions): DocxEngine;
}
