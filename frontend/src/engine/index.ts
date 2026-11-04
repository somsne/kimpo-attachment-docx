/**
 * 文档编辑引擎包入口（文档附件插件 kimpo-attachment-docx，随包文件 dist/superdoc/engine.js）。
 *
 * 单独构建为 IIFE（vite.engine.config.ts），由 plugin.umd.js 在首次进入编辑态时用 <script> 加载
 * （engine-loader.ts），加载后挂到 window.KimpoPlugins.attachment_docx_engine。vue 外部化用宿主的 window.Vue；
 * SuperDoc（AGPL-3.0）及其依赖、样式打进本文件。许可与源码地址见同目录 LICENSE / SOURCE.txt。
 */
import './polyfills';
import 'superdoc/style.css';
import './engine-theme.css';
import { createDocxEngine } from './superdoc-engine';
import type { DocxEngineModule } from '../engine-types';

/** 打进引擎包的 SuperDoc 版本（与 package.json 锁定版本一致，bundle 测试核对）。 */
export const version = '1.47.2';

export { createDocxEngine };

/** 编译期核对：本入口的导出面满足 DocxEngineModule。 */
const _shape: DocxEngineModule = { version, createDocxEngine };
void _shape;
