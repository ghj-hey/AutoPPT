# PPT Agent 2

`PPT Agent 2` 是一套可复用的离线 PPT 生成工作流，用于根据以下输入自动生成可编辑的演示文稿：

- 需求文档：`.docx` 或 `.pdf`
- 空白 PPT 模板
- 一份或多份参考 PPT

这个项目面向中文金融汇报、管理汇报、经营分析等场景，核心目标是把源材料转成**可编辑**的 PowerPoint，而不是只输出一张不可修改的图片。

## 主要特性

- 支持解析 `.docx` 和 `.pdf` 需求文档
- 可根据标题层级与正文信号动态规划页面结构
- 使用 `pptxgenjs` 生成可编辑 `.pptx`
- 支持多份参考 PPT，并提取可复用素材
- 在 Windows + Microsoft PowerPoint 环境下可导出真实渲染预览
- 保留大纲、样式、布局选项和预览等中间产物，便于检查与调整

## 工作流会做什么

整个流程会把不同输入转成统一的可编辑 PPT 产物：

- `需求文档` → 结构化内容模型
- `参考 PPT` → 可复用素材库
- `空白模板` → 页面尺寸和字体基线
- `outline + style + layout selection` → 最终 `.pptx`

同时，中间产物也会保留：

- 文档摘要
- 草稿大纲
- 草稿样式
- 页面级布局选项
- 草稿真实渲染图
- 最终真实渲染图

## 公开仓库说明

这个公开仓库只包含**代码和配置**。

不会包含：

- 实际项目素材
- 参考 PPT 内容
- 生成后的预览图和交付物
- 其他大体积二进制文件

仓库会保留目录结构，并通过 `.gitkeep` 占位，方便克隆后保持约定目录存在。

## 快速开始

### 安装依赖

```bash
npm ci
```

### 启动 Web 界面

```bash
npm run ui
```

### 生成 PPT

```bash
npm run generate -- \
  --word path/to/requirement.docx \
  --template path/to/template.pptx \
  --reference-library path/to/reference_library \
  --out path/to/output.pptx
```

### 构建或刷新素材库

```bash
npm run extract-reference-library -- \
  --input path/to/reference.pptx \
  --output path/to/reference_library
```

### 运行发布检查

```bash
npm run release-check
```

## 环境要求

- Node.js 22+
- npm
- Windows + Microsoft PowerPoint（如果需要真实渲染预览导出）

## 项目脚本

- `npm run generate`：主生成入口
- `npm run extract-reference-library`：参考 PPT 抽取流程
- `npm run ui`：启动 Web UI
- `npm test`：运行测试
- `npm run release-check`：运行公开仓库发布门禁

## 项目结构

```text
.
├── build_reference_library.js
├── report_runner.js
├── scripts/
├── src/
├── test/
├── reference_library/
├── assets/
└── README.md
```

## 语言

- [English](./README.md)
- [简体中文](./README.zh-CN.md)
