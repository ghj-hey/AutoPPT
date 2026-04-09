# PPT Agent 2

Reusable offline workflow for generating editable PPT decks from:

- a requirement document (`.docx` or `.pdf`)
- a blank PPT template
- one or more reference PPTs

This repository is designed for Chinese financial reporting, management reporting, and similar business presentation scenarios. It focuses on turning source materials into a fully editable PowerPoint workflow instead of a locked image-based export.

## Highlights

- Parse requirement documents from `.docx` and `.pdf`
- Plan slide structure dynamically from document headings and content signals
- Generate editable `.pptx` output with `pptxgenjs`
- Support multiple reference PPT uploads and reusable material extraction
- Produce real PowerPoint-rendered previews when Microsoft PowerPoint is available
- Keep intermediate artifacts visible: outline, style, layout choices, and previews

## What the workflow does

The project converts:

- `requirement doc` → structured content model
- `reference PPT(s)` → reusable material library
- `blank template` → page size and font baseline
- `outline + style + layout selection` → editable `.pptx`

It also keeps key draft artifacts accessible:

- document summary
- outline draft
- style draft
- layout options
- rendered draft preview
- final rendered preview

## Public repository scope

This public repository contains code and configuration only.

It intentionally does **not** include:

- actual project materials
- reference PPT content
- generated previews and deliverables
- other large binary assets

The folder structure is preserved with `.gitkeep` placeholders so the expected directories still exist after clone.

## Quick start

### Install

```bash
npm ci
```

### Run the web UI

```bash
npm run ui
```

### Generate a deck

```bash
npm run generate -- \
  --word path/to/requirement.docx \
  --template path/to/template.pptx \
  --reference-library path/to/reference_library \
  --out path/to/output.pptx
```

### Build or refresh the reference library

```bash
npm run extract-reference-library -- \
  --input path/to/reference.pptx \
  --output path/to/reference_library
```

### Run checks

```bash
npm run release-check
```

## Requirements

- Node.js 22+
- npm
- Microsoft PowerPoint on Windows if you want real slide preview export

## Project scripts

- `npm run generate` — main CLI generation entry
- `npm run extract-reference-library` — reference PPT extraction pipeline
- `npm run ui` — start the web UI
- `npm test` — run the test suite
- `npm run release-check` — run the public-repo release gate

## Project structure

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

## Language

- [English](./README.md)
- [简体中文](./README.zh-CN.md)
