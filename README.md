# kimpo-attachment-docx

Kimpo 附件插件：Word 文档（.docx）查看与编辑。查看用 docx-preview（Apache-2.0）分页渲染，编辑用 SuperDoc 1.47.2（AGPL-3.0，按 OOXML 结构编辑、未触碰的部件原样保留，自带修订与批注）。**本插件以 AGPL-3.0 发布。**

链路：kimpo 宿主 · 填报态（新增 / 修改 / 查看历史单据）与历史记录页打开 docx 附件 → 宿主附件外壳按 `capabilities.attachment_open.accepts` 路由到本插件 → 挂载 surface `attachment.open`，组件以 prop `ctx: AttachmentOpenContext` 收上下文。契约见主仓 `plugins/kimpo-plugin-sdk/frontend/src/attachment-open.ts`；立案 somsne/kimpo#768（分道 W）。纯前端插件，无后端。

## 许可与源码（AGPL-3.0 义务）

- 本插件全部源码以 AGPL-3.0 发布（见 `LICENSE`），源码地址：**https://github.com/somsne/kimpo-attachment-docx**。
- **向任何获得本插件的人提供完整对应源码**：插件包随附 `frontend/LICENSE`（AGPL-3.0 全文）、`frontend/SOURCE.txt`（源码地址与版本对应关系）、`frontend/superdoc/LICENSE`（SuperDoc 的 AGPL-3.0 全文）与 `frontend/superdoc/THIRD_PARTY.txt`（引擎包内第三方组件与许可）；插件界面工具栏右侧「关于」（ⓘ）显示版本、许可与源码地址。分发某个版本时，源码以与插件版本对应的 tag / 提交提供。
- SuperDoc 版本说明：锁定 **superdoc@1.47.2**（npm `legacy` 线，2026-10-08 发布，许可文本为 AGPL-3.0 全文）。**SuperDoc 2.x 起依赖专有闭源包 `@superdoc/docx-engine`**（「SuperDoc DOCX Engine Proprietary License」：禁止再分发、禁止以 AI 工具分析其实现等），不能随插件发布，故不升 2.x；升版前须先核依赖树里没有非开源组件。
- Kimpo 宿主与其他插件只经 SDK 契约与本插件交互，不受本许可影响。

### 依赖与许可

| 依赖 | 版本 | 许可 | 去向 |
|---|---|---|---|
| superdoc | 1.47.2 | AGPL-3.0 | 编辑引擎包 `superdoc/engine.js` |
| ├ SuperDoc 发布产物内联的代码：ProseMirror、Pinia、Konva、JSZip、xml-js、uuid、eventemitter3、process / buffer 垫片 | — | MIT（JSZip 为 MIT OR GPL-3.0-or-later，取 MIT） | 同上 |
| ├ SuperDoc 导入、随引擎包打入：yjs 13.6.33、y-protocols 1.0.7、lib0 0.2.119、@hocuspocus/provider 2.15.3、@hocuspocus/common 2.15.3 | — | MIT | 同上（协同编辑代码，本插件不启用） |
| docx-preview | 0.4.1 | Apache-2.0 | 主包 `plugin.umd.js` |
| jszip | 3.10.2 | MIT OR GPL-3.0-or-later（取 MIT） | 主包（守门扫描、往返摘要） |
| @kimposoft/plugin-sdk | 主仓子模块 | Apache-2.0 | 主包 |
| vue | 宿主提供 | MIT | 不打包（外部化 `window.Vue`） |
| docx | 9.12.0 | MIT | 仅生成测试夹具，不进产物 |
| jsdom 27.4.0 / happy-dom 20.9.0 / vite / vitest / vue-tsc / typescript | — | MIT / Apache-2.0 | 仅开发与测试 |

## 功能

- **查看**（docx-preview）：白纸分页页面（暗色下也是白纸），按 Word 上次排版留下的分页标记分页；工具栏 导航窗格开关 · 上一页 / 第 n / N 页 / 下一页 · 缩小 / 百分比 / 放大 / 适合宽度 · 下载（有下载权限时，经 `ctx.download()`）· 关于。左侧导航窗格「标题 / 页面」：标题按文档自带标题样式抽取（样式名 heading 1～3 或大纲级别，兼容中文文档的数字样式 ID），点击跳转。
- **守门**（`src/guard.ts`）：打开时扫描包内 `word/embeddings/`（OLE 嵌入）、`word/diagrams/`（SmartArt）、`word/charts/`（图表）、`word/vbaProject.bin`（宏），以及 `document.xml` 里的文本框（`w:txbxContent`）、形状混排（`mc:AlternateContent` / `wps:wsp`）、OLE 占位（`w:object`）、复杂域（`w:fldChar`，允许 TOC、MERGEFIELD 及目录条目自带的 PAGEREF / HYPERLINK 与页码 PAGE / NUMPAGES）。命中即只读，显示守门条幅（四语），插件内 canEdit 降为假。
- **编辑**（SuperDoc，进入编辑态才加载）：严格按设计稿画板「文档插件 · 编辑态」自绘，不用 SuperDoc 自带工具栏皮肤——
  - 功能区页签 开始 / 插入 / 布局 / 审阅 / 视图；右侧「修订」开关、查找替换、下载、关于。
  - 开始：剪贴板（粘贴 / 剪切 / 复制）、字体（字体、字号含中文字号名、增大 / 减小字号、加粗 / 斜体 / 下划线 / 删除线 / 上标 / 下标、字体颜色、突出显示、清除格式）、段落（项目符号、编号、多级列表、减少 / 增加缩进、四种对齐、行距、显示编辑标记）、样式库（文档自带样式含自定义，当前段落样式高亮）、插入（表格网格、图片）。
  - 插入：表格与表格行列操作、图片、分页符、目录、链接；布局：纸张方向 / 大小 / 页边距、分页符与分节符；审阅：修订开关、接受 / 拒绝（选区）、上一处 / 下一处、全部接受 / 全部拒绝、新建批注、审阅栏；视图：导航窗格、审阅栏、标尺、编辑标记、100% / 适合宽度。
  - 标尺；左导航窗格「标题 / 页面 / 批注」；右审阅栏列出批注与修订（接受 / 拒绝 / 回复 / 解决，时间按用户时区显示）；状态栏 第 n 页 / 共 N 页 · 字数（中文逐字计）· 语言 · 修订开关状态 · 页面视图 · 缩放。
  - 右键菜单：插件自绘四语菜单（剪切 / 复制 / 粘贴、新建批注；光标在表格里时加行列操作；选区有修订时加接受 / 拒绝）。SuperDoc 自带的英文右键菜单、浮动批注气泡、右侧批注栏都关闭。
  - 查找替换：Ctrl+F / Ctrl+H 或功能区按钮，浮动面板，支持区分大小写、逐个替换与全部替换。
- **编辑流**（`src/edit-flow.ts`，照搬文本插件，改动须两边同步）：外壳切到编辑且 canEdit → 先 `ctx.beginEdit()` 取锁，取到锁才加载编辑引擎；有改动 → `setDirty(true)`（按「内容指纹」判断：正文 ProseMirror 文档 + 批注签名与基线比对，载入期规范化不算改动）；宿主「保存」经 `ctx.onSaveRequest` 回到插件；EDIT_LOCK_LOST 保留内容、可「重新获取并保存」或下载副本。
- **往返自检**（`src/docx-summary.ts`）：
  - 导入自检：引擎就绪后比对原文件摘要与编辑器模型，不一致（编辑器没装下原文件）转只读并出条幅；
  - 保存自检：导出 docx 后重新解析，与编辑器模型比对 段落数、表格数、正文文字（去空白）三项，一致才 `session.save`（文件名与类型沿用原文件），否则拒存、上报 warn 日志、出条幅，修改保留。
  - 摘要只看 `document.xml` 正文；修订里的删除文字不计、插入计入；`mc:Fallback` 备份与纵向 / 横向合并续格不计。
- 安全与隐私：SuperDoc 默认开启遥测（向 `ingest.superdoc.dev` 发数据），插件显式关闭（`telemetry: { enabled: false }`，manifest 声明无网络权限）。插件不屏蔽右键、选中与快捷键；水印由宿主叠加。
- 四语（zh-CN / zh-TW / en-US / ja-JP），颜色取宿主主题变量，暗色随 `html.dark`，页面始终白纸。

## 打包与随包结构

```
frontend/dist/
├── plugin.umd.js          # 主包：查看、守门、编辑态界面；约 326 KB（gzip 约 95 KB）
├── LICENSE                # 本插件 AGPL-3.0 全文
├── SOURCE.txt             # 源码地址
└── superdoc/
    ├── engine.js          # 编辑引擎包（SuperDoc 1.47.2 + 封装）；约 6.5 MB（gzip 约 1.9 MB），进入编辑态才加载
    ├── LICENSE            # SuperDoc AGPL-3.0 全文
    └── THIRD_PARTY.txt    # 引擎包内第三方组件与许可
```

- manifest 入口写包内 canonical 路径 `frontend/plugin.umd.js`。打进 KPP 时**整个 `frontend/dist/` 的内容收进 KPP 的 `frontend/`**，保持 `superdoc/` 子目录与 `plugin.umd.js` 同级（与 PDF 插件的 `pdfjs/` 同一做法）。插件从自身脚本地址（`document.currentScript.src`，模块求值时取一次）推出 `<脚本目录>/superdoc/engine.js`，首次进入编辑态时插 `<script>` 加载，挂到 `window.KimpoPlugins.attachment_docx_engine`。
- dev 源码直服（宿主 `resolveDevFrontendAsset`）目前只把入口同目录文件映射到 `frontend/dist/`，`superdoc/` 子目录待主仓扩展；扩展前 dev 环境里进入编辑态会提示「编辑器加载失败」。
- 引擎包构建（`vite.engine.config.ts`）：IIFE、vue 外部化、目标 chrome88；SuperDoc 发布产物里内联的 `process` 浏览器垫片让 91 处 `NODE_ENV` 判断走开发分支（会让 Pinia 往页面全局塞调试队列并崩溃），构建时定为 production。Chromium 88 缺的 `structuredClone`、`crypto.randomUUID`（内网 http 部署时新浏览器也没有）、`Array/String.prototype.at`、`findLast`、`Object.hasOwn` 由 `src/engine/polyfills.ts` 在缺失时补到页面全局（只随引擎包加载）。
- 浏览器下限 Chromium 88：只做了构建目标与垫片，**尚未在真实 Chromium 88 内核浏览器上验证**。不用 SharedArrayBuffer / WASM 线程（产物里出现的只是第三方库的特性探测）。

## PoC 往返验收

```bash
cd frontend
pnpm fixtures                                   # 重新生成 fixtures/roundtrip/*.docx（docx 包，MIT）
node scripts/roundtrip.mjs <样件目录>            # 逐件：段落 / 表格 / 图片 / 文字 原→导出，退出码 0 = 全部零差异
node scripts/roundtrip.mjs <样件目录> --out <目录>  # 另把导出结果写出来，供在 Word 里人工比对版式与页数
node scripts/roundtrip.mjs <样件目录> --json     # 机器可读
```

导入与导出用与插件编辑态同一版本的 SuperDoc（headless + jsdom）。页数差需要排版引擎，脚本不测，需在 Word 里人工比对。仓内 8 份夹具（标题样式、编号与多级列表、表格含合并与嵌套、页眉页脚与页码、图片、修订、批注、自定义样式与分节）全部零差异。

## 开发

```bash
cd frontend
pnpm install
pnpm test          # vitest：查看 / 守门 / 编辑保存 / 自检拒存 / 导入自检 / 失锁 / 修订批注 / 右键菜单 / 引擎封装 / 8 份夹具真实往返
pnpm typecheck     # vue-tsc
pnpm build         # 主包 + 引擎包
pnpm test:bundle   # 构建后的产物断言（node --test）
KAD_ENGINE_DEBUG=1 npx vite build -c vite.engine.config.ts   # 排查引擎问题时出未压缩引擎包
```

- `@kimposoft/plugin-sdk` 暂以 `file:` 指向主仓子模块 `plugins/kimpo-plugin-sdk/frontend`，所以本仓需在主仓 `plugins/` 下构建。
- manifest 预检：主仓 SDK 目录下 `GOWORK=off go run ./cmd/kimpo-plugin-sign lint -official <解包目录>`（解包目录 = manifest.json + frontend/ 下的 dist 内容与 locales）。

## 已知限制（一期）

- 「粘贴」按钮只能粘纯文本（浏览器只开放剪贴板纯文本读取，且需用户授权、非 https 不可用）；带格式粘贴用 Ctrl+V。
- 「多级列表」：SuperDoc 1.x 没有多级列表样式库，按钮先建编号列表，层级用「增加 / 减少缩进」或 Tab 调整。
- 页面设置作用于全文各节（与 Word「应用于整篇文档」一致），没有逐节设置；页面设置、页眉页脚、表格属性没有做成对话框（页面设置为下拉菜单；页眉页脚在页面上双击编辑；表格属性未做）。
- 批注署名：契约（AttachmentOpenContext）暂无当前用户名，插件按可选字段 `env.userDisplayName` / `env.userName` / `user.displayName` 读取，缺省署名「未署名」，待宿主补字段。
- 查看态不显示批注与修订标记（docx-preview 这两项是实验功能，未开）；没有 Word 分页标记的文档（非 Word 生成）查看态按节分页，页数可能少于 Word；编辑态由 SuperDoc 自己排版分页。
- 功能区宽度不足时不折叠为下拉（只在外壳宽度不足时默认收起两侧窗格）；方案「其他」里的「自动保存草稿到浏览器本地」本期未做。
