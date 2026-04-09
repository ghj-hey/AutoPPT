# PPT Agent 2

![CI](https://github.com/ghj-hey/AutoPPT/actions/workflows/ci.yml/badge.svg)
![License](https://img.shields.io/github/license/ghj-hey/AutoPPT)
![Node](https://img.shields.io/badge/node-22%2B-339933)

PPT Agent 2 is an offline PowerPoint workflow for turning requirements docs and reference decks into editable, business-ready presentations.

It is built for Chinese financial reporting, management reporting, and similar presentation-heavy workflows where editable output matters more than a flat image export.

## At a glance

- Parse requirement documents from `.docx` and `.pdf`
- Plan slide structure from headings, tables, images, and content signals
- Generate editable `.pptx` files with `pptxgenjs`
- Reuse materials extracted from one or more reference PPTs
- Export real PowerPoint-rendered slide previews when PowerPoint is available
- Keep the drafting pipeline visible: summary, outline, style, layout, and preview artifacts

## Project status

- Public repository: code and configuration only
- CI: GitHub Actions running on every push and pull request
- Release gate: `npm run release-check`
- Runtime: Node.js 22+

## Public repository scope

This GitHub repository intentionally contains **code and configuration only**.

It does not include:

- actual project materials
- reference deck content
- generated previews or deliverables
- other large binary assets

The expected directory structure is preserved with `.gitkeep` placeholders so the repository still clones cleanly and remains easy to understand.

## How the workflow fits together

The project turns source inputs into a fully editable deck:

- `requirement doc` → structured content model
- `reference PPT(s)` → reusable material library
- `blank template` → page size and font baseline
- `outline + style + layout selection` → final `.pptx`

Intermediate artifacts stay available for inspection and correction:

- document summary
- outline draft
- style draft
- layout options
- rendered draft preview
- final rendered preview

## Quick start

### 1. Install dependencies

```bash
npm ci
```

### 2. Start the web UI

```bash
npm run ui
```

### 3. Generate a deck

```bash
npm run generate -- \
  --word path/to/requirement.docx \
  --template path/to/template.pptx \
  --reference-library path/to/reference_library \
  --out path/to/output.pptx
```

### 4. Refresh the reference library

```bash
npm run extract-reference-library -- \
  --input path/to/reference.pptx \
  --output path/to/reference_library
```

### 5. Run the release gate

```bash
npm run release-check
```

## Requirements

- Node.js 22+
- npm
- Microsoft PowerPoint on Windows for real slide preview export

## Scripts

- `npm run generate` — main CLI generation entry
- `npm run extract-reference-library` — reference PPT extraction pipeline
- `npm run ui` — start the web UI
- `npm test` — run the test suite
- `npm run release-check` — run the public-repo release gate

## Repository layout

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

## Languages

- English
- 简体中文

See [README.zh-CN.md](./README.zh-CN.md) for the Chinese version.
