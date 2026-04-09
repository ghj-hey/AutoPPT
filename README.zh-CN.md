# PPT Agent 2

[English](./README.md) | [简体中文](./README.zh-CN.md)

`PPT Agent 2` 是一套可复用的离线 PPT 生成工作流，用于根据以下输入自动生成可编辑的银行金融风格 PPT：

- 需求文档：`.docx` 或 `.pdf`
- 空白 PPT 模板
- 一份或多份参考 PPT / PPT 素材

这套系统面向中文金融汇报、经营分析、管理汇报等场景设计，但生成逻辑已经不再绑定单一业务主题。当前版本会根据上传文档的章节结构、表格、图片和参考风格动态规划页面，并同时使用：

- 本次上传的参考素材
- 已累计沉淀的素材库
- 当前选择的布局模板库

## 这个项目解决什么问题

整个流程会把不同输入转成统一的可编辑 PPT 产物：

- `需求文档` -> 结构化内容模型
- `参考 PPT / 素材 PPT` -> 可复用素材库
- `空白模板` -> 页面尺寸、版心和字体基线
- `outline + style + layout selection` -> 最终可编辑 `.pptx`

同时，中间产物不会被隐藏，而是尽量可视化、可检查、可校正：

- 文档摘要
- 草稿大纲
- 草稿样式
- 页面级布局选项
- 草稿真实渲染图
- 最终真实渲染图

## 核心能力

- 支持 `docx` 和 `pdf` 作为需求文档输入。
- 能识别中文标题层级，例如 `一、`、`（一）`、`1`、`1.1`、`1.1.1`，以及 Word 大纲层级。
- 支持动态规划 `2` 到 `10` 页 PPT，而不是固定某一种业务模板。
- 同一会话支持上传多份参考 PPT / 素材 PPT。
- 会从新上传的 PPT 中抽取素材，并仅将新增素材合并进主素材库。
- 生成时同时利用“当前上传素材”和“已有素材库”。
- 可通过 PowerPoint COM 导出真实页级截图，而不是只给结构化缩略图。
- 支持按页面类型切换不同布局模板。
- 保留 `JSON 校正` 能力，可直接修改 `outline` 和 `style`。
- 支持接入语义模型辅助文档理解、布局建议和最终效果复核。

## 当前生成模型

系统使用的是“通用页面类型”，而不是单一业务故事线。当前主页面类型包括：

- `cover`
- `summary_cards`
- `table_analysis`
- `process_flow`
- `bullet_columns`
- `image_story`
- `action_plan`
- `key_takeaways`

页面规划会综合以下信号进行判断：

- 标题层级和章节结构
- 文字密度
- 表格数量和表格规模
- 图片数量和长宽比
- 流程类关键词
- 指标 / 结果 / 成效类关键词
- 下一步 / 行动计划类关键词

这意味着：即使参考风格不变，换一份需求文档，也会生成不同的 PPT 结构。

## 项目结构

```text
E:\codex\ppt_agent2
|-- report_runner.js
|-- build_reference_library.js
|-- package.json
|-- README.md
|-- README.zh-CN.md
|-- reference_library
|   |-- work_meeting
|   |   |-- layout_templates.json
|   |   `-- reusable_materials.json
|   |-- master
|   `-- libraries
|-- src
|   |-- api
|   |   `-- server.js
|   |-- cli
|   |   |-- build_reference_library.js
|   |   `-- report_runner.js
|   |-- services
|   |   |-- deckRendererService.js
|   |   |-- documentParserService.js
|   |   |-- outlinePlannerService.js
|   |   |-- powerPointPreviewService.js
|   |   |-- referenceLibraryService.js
|   |   |-- workflowService.js
|   |   `-- workflowSessionService.js
|   |-- ui
|   |   `-- public
|   |       |-- app.js
|   |       |-- index.html
|   |       `-- styles.css
|   `-- utils
|       |-- fileUtils.js
|       |-- hashUtils.js
|       `-- pathConfig.js
`-- workspace
    |-- sessions
    `-- uploads
```

## 模块说明

### `report_runner.js`

根目录下的轻量入口，转发到 [src/cli/report_runner.js](/E:/codex/ppt_agent2/src/cli/report_runner.js)。

### `build_reference_library.js`

根目录下的轻量入口，转发到 [src/cli/build_reference_library.js](/E:/codex/ppt_agent2/src/cli/build_reference_library.js)。

### `src/cli/report_runner.js`

CLI 主生成入口。

主要职责：

- 解析命令行参数
- 加载空白 PPT 模板
- 加载生效中的参考素材库
- 加载布局模板库
- 解析需求文档
- 生成 `outline.json`
- 生成 `style.json`
- 生成 `document_summary.json`
- 解析页面级模板分配
- 渲染最终 `.pptx`

重要参数：

- `--word` 或 `--doc`
- `--template`
- `--reference-library`
- `--layout-library`
- `--layout-set`
- `--pages auto|2..10`
- `--out`

### `src/cli/build_reference_library.js`

参考 PPT 抽取流程。

主要职责：

- 解包 PPT 内媒体资源
- 导出参考 PPT 的真实页面预览
- 按类别归档媒体资源
- 汇总颜色、字体、形状、表格风格和页面级信号
- 生成：
  - `catalog.json`
  - `reusable_materials.json`
  - `media/`
  - `categorized/`
  - `previews/`

### `src/services/documentParserService.js`

需求文档解析器。

主要职责：

- 解析 `.docx`
- 解析 `.pdf`
- 抽取段落节点
- 抽取表格
- 抽取内嵌图片
- 读取模板 PPT 的页面尺寸和主题字体
- 生成统一格式的文档摘要

### `src/services/outlinePlannerService.js`

动态页面规划器。

主要职责：

- 根据中文编号和大纲元数据识别标题层级
- 对候选章节进行打分
- 在 `pages=auto` 时推断合理页数
- 将文档内容映射到通用页面类型
- 为每一页生成备注说明

这是项目从“单一业务模板”演进到“文档驱动可复用方案”的核心模块之一。

### `src/services/deckRendererService.js`

基于 `pptxgenjs` 的可编辑 PPT 渲染器。

主要职责：

- 渲染各类页面
- 从素材库中选择图标和视觉素材
- 按页面类型应用不同布局变体
- 在必要时自动缩小字体以适配文本
- 绘制表格、卡片、标题签、徽标、柱条、面板和图片区块

### `src/services/powerPointPreviewService.js`

真实预览导出链路。

主要职责：

- 通过 PowerPoint COM 打开生成的 PPT
- 将每一页导出为 PNG
- 返回给 UI 用于展示的页级预览

这条链路比“结构化缩略图”更可信，因为它使用的是 PowerPoint 自身的渲染引擎。

### `src/services/referenceLibraryService.js`

参考素材库生命周期管理器。

主要职责：

- 列出可用素材库
- 从新上传的参考 PPT 中抽取新库
- 将新增素材合并到主素材库
- 将多份素材库组合成会话生效素材库
- 按文件哈希去重
- 在前端展示时按来源名去重

### `src/services/workflowService.js`

UI 对应的工作流编排层。

主要职责：

- 归档上传文件
- 解析需求文档
- 抽取或合并上传的参考 PPT
- 生成草稿大纲和样式
- 计算页面级布局选择
- 渲染草稿 PPT
- 导出草稿真实渲染图
- 在人工修正后重新生成最终 PPT
- 导出最终真实渲染图

### `src/services/workflowSessionService.js`

会话目录管理器。

主要职责：

- 创建会话目录
- 持久化 `session.json`
- 归档上传文件
- 解析会话输入输出路径

### `src/api/server.js`

本地 UI 使用的 Express API。

当前主要路由：

- `GET /api/health`
- `GET /api/libraries`
- `POST /api/workflow/create`
- `POST /api/workflow/generate`
- `GET /api/workflow/:sessionId`
- `GET /api/download/:sessionId/:fileType`

### `src/ui/public/index.html`

本地工作台主页面。

主要功能：

- 上传参考素材、模板和需求文档
- 选择生成页数
- 选择布局方案集
- 查看流程阶段和文档摘要
- 切换页面级布局模板
- 查看并修改 `outline/style` 的原始 JSON
- 下载草稿和最终产物

### `src/ui/public/app.js`

浏览器端逻辑。

主要职责：

- 提交上传请求
- 渲染预览卡片
- 渲染布局选择器
- 渲染 JSON 编辑器
- 将 JSON 修改同步回工作流模型
- 通过 `blob` 方式触发下载

### `src/ui/public/styles.css`

工作台视觉层。

主要职责：

- 定义本地 UI 的银行金融风格表现
- 适配桌面与移动端下的上传、预览和编辑网格

### `src/utils/fileUtils.js`

JSON 读写、目录创建、文件复制和目录枚举等通用工具。

### `src/utils/hashUtils.js`

素材去重时使用的哈希工具。

### `src/utils/pathConfig.js`

集中管理路径配置：

- 工作区根目录
- 会话根目录
- 上传根目录
- 默认参考库
- 默认布局库
- 主素材库
- 抽取素材库存放根目录

## 素材库模型

素材系统是分层设计的，因此随着更多参考 PPT 的加入，整个项目的素材库会不断丰富。

主要产物：

- `media/`：原始抽取媒体
- `categorized/`：分类后的媒体
- `catalog.json`：统计信息和抽取摘要
- `reusable_materials.json`：可复用预设与素材集合

当前重点跟踪的可复用对象包括：

- 配色 token
- 图标素材
- 品牌素材
- 插画素材
- 截图素材
- 组件预设
- 素材标签
- 用途标签

示例素材标签：

- `branding`
- `vector`
- `bitmap`
- `icon`
- `screenshot`
- `small`
- `wide`
- `reusable`

示例用途标签：

- `header-logo`
- `summary-card`
- `process-flow`
- `data-analysis`
- `system-preview`

## 布局模板模型

默认布局模板库位于：

- `reference_library/work_meeting/layout_templates.json`

它把两个概念拆开了：

- `layout`：内容放在哪里
- `material`：内容看起来像什么

当前支持的布局方案集包括：

- `bank_finance_default`
- `bank_finance_dense`
- `bank_finance_visual`
- `bank_finance_highlight`
- `bank_finance_boardroom`
- `bank_finance_reporting`

每种页面类型都有多套可互换模板。例如：

- `table_analysis`：`split`、`dense`、`highlight`、`visual`
- `process_flow`：`three_lane`、`cards`、`ladder`
- `bullet_columns`：`dual`、`triple`、`staggered`
- `image_story`：`split`、`focus`、`gallery`
- `action_plan`：`timeline`、`matrix`、`stacked`

## 工作流

### 第 1 步：上传输入

输入包括：

- 一份需求文档
- 一份空白 PPT 模板
- 零份或多份参考 PPT / 素材 PPT

### 第 2 步：解析文档

解析器会提取：

- 正文段落
- 标题层级
- 表格
- 图片
- 模板尺寸与字体

### 第 3 步：构建有效风格上下文

系统会组合：

- 本次上传参考 PPT 的抽取结果
- 当前选中的基础素材库
- 已累计沉淀的主素材库

### 第 4 步：规划大纲

大纲规划器会：

- 决定页数
- 选择页面类型
- 分配章节内容
- 形成草稿页面序列

### 第 5 步：渲染草稿

渲染器会生成：

- 草稿 PPT
- 草稿 Outline
- 草稿 Style
- 草稿 Notes
- 草稿真实渲染截图

### 第 6 步：人工校正

操作人可以：

- 切换布局方案集
- 切换单页布局模板
- 查看草稿预览和阶段摘要
- 在需要时直接修改 `outline.json` 和 `style.json`

### 第 7 步：最终渲染

系统会重新生成：

- 最终 PPT
- 最终 Outline
- 最终 Style
- 最终布局清单
- 最终 Notes
- 最终真实渲染截图

## 常用命令

安装依赖：

```bash
npm install
```

启动本地 UI：

```bash
npm run ui
```

通过 CLI 生成 PPT：

```bash
npm run generate -- --doc "E:\\path\\input.docx" --template "E:\\path\\template.pptx" --pages auto --out "E:\\path\\output"
```

抽取新的参考素材库：

```bash
npm run extract-reference-library -- --ppt "E:\\path\\reference.pptx" --out "E:\\path\\library"
```

## 会话输出目录

每一次 UI 工作流会话首先会写入：

- `workspace/sessions/<sessionId>/input`
- `workspace/sessions/<sessionId>/intermediate`
- `workspace/sessions/<sessionId>/output`

在最终 PPT 生成后，系统会把交付文件迁移到
`workspace/deliverables/<sessionId>/output`，并删除原始会话工作区，
包括上传缓存和中间产物。`reference_library/` 里的持久化素材库
会继续保留。

典型中间文件：

- `outline.json`
- `style.json`
- `page_notes.md`
- `document_summary.json`
- `reference_summary.json`
- `layout_options.json`
- `draft_generated.pptx`
- `rendered_preview/`

典型归档后最终文件：

- `workflow_generated.pptx`
- `outline.final.json`
- `style.final.json`
- `document_summary.final.json`
- `document_structure.final.json`
- `document_structure.final.md`
- `layout.final.json`
- `page_notes.final.md`
- `rendered_preview/`

## 布局中文标签与多样性验证

当前棕地布局契约保持内部模板 id 稳定，同时保证所有用户可见标签为中文：

- UI 下拉框展示中文 `displayName` / `label`
- `layout_options.json` 在保留 `id` / `currentTemplate` /
  `recommendedTemplate` 等机器字段稳定的同时输出中文显示字段
- 下载清单与归档清单保持中文标题，同时继续使用稳定的机器 `type` 字段

推荐本地验证命令：

```bash
node --check src/ui/public/app.js
node --check src/services/layoutSelectionService.js
node --check src/services/workflowService.js
node --check src/services/deckRendererService.js
node --test test/layoutVerification.test.js
```

## 当前验证结论

当前实现已经支持：

- 非三农类文档的动态规划
- 类似 `一、`、`1.1`、`1.1.1` 的层级标题自动处理
- 多份参考 PPT 的抽取和合并
- 通过 PowerPoint 真实渲染导出页级截图
- 按页面类型复用布局切换

## 已知约束

- 真实渲染截图依赖本机安装 Microsoft PowerPoint。
- 当前默认风格主要优化的是中文银行 / 金融汇报场景，若是完全不同的视觉风格，通常需要另一套默认布局库。
- 自动文本适配已经做过增强，但对于极高密度原始文档，仍可能需要手动切换布局或提高页数。

## 推荐使用方式

- 实际生产建议优先使用 UI 工作台。
- 批量生成或排查问题时建议用 CLI。
- 持续向主素材库补充高质量参考 PPT，可以让后续生成结果逐步变好。
- 当需求文档信息密度较高时，建议先用 `auto`，如果草稿仍然拥挤，再提高到 `7` 到 `10` 页。
