# PPT Agent 2

![CI](https://github.com/ghj-hey/AutoPPT/actions/workflows/ci.yml/badge.svg)
![License](https://img.shields.io/github/license/ghj-hey/AutoPPT)
![Node](https://img.shields.io/badge/node-22%2B-339933)

`PPT Agent 2` 是一套离线 PowerPoint 工作流，用于把需求文档和参考 PPT 转成**可编辑**、可复用的商务演示文稿。

它主要面向中文金融汇报、管理汇报、经营分析等场景，强调的是“可编辑的 PPT 产物”，而不是只输出平面图片。

## 一览

- 支持解析 `.docx` 和 `.pdf` 需求文档
- 能根据标题、表格、图片和内容信号动态规划页面结构
- 使用 `pptxgenjs` 生成可编辑 `.pptx`
- 支持从多个参考 PPT 中抽取并复用素材
- 在可用时可通过 Microsoft PowerPoint 导出真实页面预览
- 保留摘要、大纲、样式、布局和预览等中间产物，方便检查与修正

## 项目状态

- 公开仓库：仅包含代码和配置
- CI：GitHub Actions 在每次 push 和 pull request 时运行
- 发布门禁：`npm run release-check`
- 运行环境：Node.js 22+

## 公开仓库范围

这个 GitHub 仓库只包含**代码和配置**。

不会包含：

- 实际项目素材
- 参考 PPT 内容
- 生成后的预览图和交付物
- 其他大体积二进制文件

仓库保留了必要的目录结构，并通过 `.gitkeep` 占位，方便克隆后保持目录完整、结构清晰。

## 工作流如何串起来

系统会把不同输入转成统一的可编辑 PPT 产物：

- `需求文档` → 结构化内容模型
- `参考 PPT` → 可复用素材库
- `空白模板` → 页面尺寸和字体基线
- `outline + style + layout selection` → 最终 `.pptx`

同时，以下中间产物也会保留，便于校验和调整：

- 文档摘要
- 草稿大纲
- 草稿样式
- 页面级布局选项
- 草稿真实渲染图
- 最终真实渲染图

## 快速开始

### 1. 安装依赖

```bash
npm ci
```

### 2. 启动 Web 界面

```bash
npm run ui
```

### 3. 生成 PPT

```bash
npm run generate -- \
  --word path/to/requirement.docx \
  --template path/to/template.pptx \
  --reference-library path/to/reference_library \
  --out path/to/output.pptx
```

### 4. 刷新素材库

```bash
npm run extract-reference-library -- \
  --input path/to/reference.pptx \
  --output path/to/reference_library
```

### 5. 运行发布检查

```bash
npm run release-check
```

## 环境要求

- Node.js 22+
- npm
- Windows + Microsoft PowerPoint（如果需要真实页面预览导出）

## 脚本

- `npm run generate`：主生成入口
- `npm run extract-reference-library`：参考 PPT 抽取流程
- `npm run ui`：启动 Web UI
- `npm test`：运行测试
- `npm run release-check`：运行公开仓库发布门禁

## 仓库结构

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

- English
- 简体中文

中文版本见 [README.zh-CN.md](./README.zh-CN.md)。
