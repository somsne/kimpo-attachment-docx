/**
 * 文档附件插件（kimpo-attachment-docx）前端 UMD 入口。
 *
 * 链路：kimpo 宿主 · 填报态 / 历史记录页打开 docx 附件 → 宿主附件外壳按 manifest
 * capabilities.attachment_open.accepts 路由到本插件 → 调 window.KimpoPlugins.attachment_docx.createSurface(input)
 * → 挂载 DocxStage，组件以 prop `ctx: AttachmentOpenContext` 收上下文。
 * hostBridge 一并作为 prop 传入，用于多语言（i18n）与确认框（ui.confirm）。
 * 宿主标题栏「保存」经 ctx.onSaveRequest 回到插件（进入编辑态后登记，见 edit-flow.ts）。
 * 编辑引擎（SuperDoc，AGPL-3.0）不在本文件：随包 superdoc/engine.js，进入编辑态才加载（engine-loader.ts）。
 */
import { defineAttachmentOpenSurface } from '@kimposoft/plugin-sdk/attachment-open';
import './styles.css';
import DocxStage from './DocxStage.vue';

export const createSurface = defineAttachmentOpenSurface(input => ({
  component: DocxStage,
  props: { hostBridge: input.hostBridge }
}));
