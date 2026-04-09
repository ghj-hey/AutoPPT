# PPT Agent 2

[English](./README.md) | [简体中文](./README.zh-CN.md)

Reusable offline workflow for generating editable bank-style PPT decks from:

- a requirement document (`.docx` or `.pdf`)
- a blank PPT template
- one or more reference PPTs

The project is designed for Chinese financial and management-reporting scenarios, but the generation logic is no longer tied to a single business topic. It now plans slides dynamically from the uploaded document structure and blends styles from both the accumulated material library and the current uploaded references.

## What This Project Solves

The workflow converts:

- `requirement doc` -> structured content model
- `reference PPT(s)` -> reusable material library
- `blank template` -> page size and font baseline
- `outline + style + layout selection` -> editable `.pptx`

It also keeps the middle steps visible and editable:

- document summary
- outline draft
- style draft
- layout options
- rendered draft preview
- final rendered preview

## Key Capabilities

- Supports `docx` and `pdf` requirement documents.
- Detects Chinese heading structures such as `一、`, `（一）`, `1`, `1.1`, `1.1.1`, plus Word outline levels.
- Dynamically plans `2` to `10` slides instead of using one fixed business deck.
- Supports multiple reference PPT uploads in the same session.
- Extracts new materials from uploaded references and merges only new assets into the master library.
- Uses both `current uploaded references` and `existing material libraries` during generation.
- Provides real PowerPoint-rendered page previews through PowerPoint COM export, not only structural mock previews.
- Supports page-level layout switching by slide type.
- Keeps draft and final artifacts transparent, while still allowing raw JSON editing when needed.

## Current Generation Model

The generator works with generic page types, not a single business storyline:

- `cover`
- `summary_cards`
- `table_analysis`
- `process_flow`
- `bullet_columns`
- `image_story`
- `action_plan`
- `key_takeaways`

The planner chooses slide types according to document signals such as:

- heading hierarchy
- sentence density
- table count
- image count
- process keywords
- result and indicator keywords
- action-plan keywords

This means a new uploaded document can produce a different deck structure even when the reference style remains the same.

## Project Structure

```text
E:\codex\ppt_agent2
|-- report_runner.js
|-- build_reference_library.js
|-- package.json
|-- README.md
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

## Module Guide

### `report_runner.js`

Thin root wrapper. It forwards execution to `src/cli/report_runner.js`.

### `build_reference_library.js`

Thin root wrapper. It forwards execution to `src/cli/build_reference_library.js`.

### `src/cli/report_runner.js`

Main CLI generator entry.

Responsibilities:

- parse CLI arguments
- load the blank PPT template
- load the effective reference library
- load the layout template library
- parse the requirement document
- build `outline.json`
- build `style.json`
- build `document_summary.json`
- resolve page-level template assignments
- render the final `.pptx`

Important inputs:

- `--word` or `--doc`
- `--template`
- `--reference-library`
- `--layout-library`
- `--layout-set`
- `--pages auto|2..10`
- `--out`

### `src/cli/build_reference_library.js`

Reference PPT extraction pipeline.

Responsibilities:

- unpack PPT media
- export real page previews from the reference PPT
- classify media into reusable buckets
- summarize colors, fonts, shapes, table styles, and page-level signals
- generate:
  - `catalog.json`
  - `reusable_materials.json`
  - `media/`
  - `categorized/`
  - `previews/`

### `src/services/documentParserService.js`

Requirement-document parser.

Responsibilities:

- parse `.docx`
- parse `.pdf`
- extract paragraph nodes
- extract tables
- extract embedded images
- read PPT template size and theme fonts
- build normalized document summaries

### `src/services/outlinePlannerService.js`

Dynamic slide planner.

Responsibilities:

- detect heading hierarchy from Chinese numbering and outline metadata
- score candidate sections
- infer a reasonable slide count when `pages=auto`
- map document content into generic page types
- build notes for each slide

This is the main place where the project shifted from a single business layout to reusable document-driven planning.

### `src/services/deckRendererService.js`

Editable PPT renderer based on `pptxgenjs`.

Responsibilities:

- render each page type
- pick icons and visual assets from the material library
- apply layout variants by page type
- fit text by shrinking font sizes when needed
- draw tables, cards, ribbons, badges, bars, panels, and image blocks

### `src/services/powerPointPreviewService.js`

Real preview export chain.

Responsibilities:

- open generated PPT in PowerPoint via COM
- export each slide as PNG
- return page-level `dataUrl` previews for the UI

This is the actual rendered preview path. It is more trustworthy than a synthetic structural thumbnail because it uses PowerPoint's own rendering engine.

### `src/services/referenceLibraryService.js`

Reference-library lifecycle manager.

Responsibilities:

- list available libraries
- extract a library from a newly uploaded reference PPT
- merge new assets into the master library
- compose multiple libraries into one effective session library
- deduplicate assets by file hash
- deduplicate duplicate library entries by source name when presenting library choices

### `src/services/workflowService.js`

Session orchestration layer for the UI.

Responsibilities:

- stage uploads
- parse requirement documents
- extract or merge uploaded reference PPTs
- build draft outline and style
- compute page-level layout choices
- render draft PPT
- export real draft previews
- regenerate final PPT after human corrections
- export real final previews

### `src/services/workflowSessionService.js`

Session directory manager.

Responsibilities:

- create session directories
- persist `session.json`
- stage uploaded files
- resolve session input and output paths

### `src/api/server.js`

Express API used by the local UI.

Routes:

- `GET /api/health`
- `GET /api/libraries`
- `POST /api/workflow/create`
- `POST /api/workflow/generate`
- `GET /api/workflow/:sessionId`
- `GET /api/download/:sessionId/:fileType`

### `src/ui/public/index.html`

Main local workbench page.

Functions:

- upload references, template, and requirement doc
- choose page count
- choose layout set
- inspect summaries and workflow stages
- switch page layouts
- review and adjust raw outline/style JSON when needed
- download draft and final artifacts

### `src/ui/public/app.js`

Browser-side logic.

Responsibilities:

- submit uploads
- render preview cards
- render layout selectors
- render JSON editors
- synchronize JSON edits back into the workflow model
- trigger blob-based downloads

### `src/ui/public/styles.css`

Workbench visual layer.

Responsibilities:

- bank-style visual presentation for the local app
- responsive upload, preview, and editing grids

### `src/utils/fileUtils.js`

Shared helpers for JSON I/O, directory creation, copying, and directory listing.

### `src/utils/hashUtils.js`

Hash helpers used for material deduplication.

### `src/utils/pathConfig.js`

Central path registry for:

- workspace roots
- session root
- upload root
- default reference library
- default layout library
- master library
- extracted library root

## Material Library Model

The material system is layered so that new reference PPTs can keep enriching the project over time.

Main outputs:

- `media/`: raw extracted media
- `categorized/`: grouped media by category
- `catalog.json`: statistics and extraction summary
- `reusable_materials.json`: reusable presets and asset collections

The reusable library currently tracks:

- palette tokens
- icon assets
- branding assets
- illustration assets
- screenshot assets
- component presets
- material tags
- usage tags

Example material tags:

- `branding`
- `vector`
- `bitmap`
- `icon`
- `screenshot`
- `small`
- `wide`
- `reusable`

Example usage tags:

- `header-logo`
- `summary-card`
- `process-flow`
- `data-analysis`
- `system-preview`

## Layout Template Model

The default layout library lives at:

- `reference_library/work_meeting/layout_templates.json`

It separates:

- `layout`: where content goes
- `material`: what the content looks like

Supported layout sets currently include:

- `bank_finance_default`
- `bank_finance_dense`
- `bank_finance_visual`
- `bank_finance_highlight`
- `bank_finance_boardroom`
- `bank_finance_reporting`

Each slide type has multiple interchangeable templates. Example:

- `table_analysis`: split, dense, highlight, visual
- `process_flow`: three-lane, cards, ladder
- `bullet_columns`: dual, triple, staggered
- `image_story`: split, focus, gallery
- `action_plan`: timeline, matrix, stacked

## Workflow

### Step 1: Upload

Inputs:

- one requirement document
- one blank PPT template
- zero or more reference PPTs

### Step 2: Parse

The parser extracts:

- paragraphs
- heading structure
- tables
- images
- template size and fonts

### Step 3: Build effective style context

The system combines:

- current uploaded reference PPT extractions
- the selected base library
- the accumulated master library

### Step 4: Plan outline

The outline planner:

- decides page count
- chooses page types
- assigns sections
- creates a draft slide sequence

### Step 5: Render draft

The renderer produces:

- draft PPT
- draft outline
- draft style
- draft notes
- real rendered slide previews

### Step 6: Human correction

The operator can:

- switch layout set
- switch page-level layout templates
- review draft previews and workflow summaries
- edit `outline.json` and `style.json` directly when needed
- regenerate the final deck from corrected JSON

### Step 7: Final render

The system regenerates:

- final PPT
- final outline
- final style
- final layout manifest
- final notes
- real final slide previews

## Commands

Install dependencies:

```bash
npm install
```

Start the local UI:

```bash
npm run ui
```

Generate a deck from CLI:

```bash
npm run generate -- --doc "E:\\path\\input.docx" --template "E:\\path\\template.pptx" --pages auto --out "E:\\path\\output"
```

Extract a new reference library:

```bash
npm run extract-reference-library -- --ppt "E:\\path\\reference.pptx" --out "E:\\path\\library"
```

## Session Output Layout

Every UI workflow session first writes to:

- `workspace/sessions/<sessionId>/input`
- `workspace/sessions/<sessionId>/intermediate`
- `workspace/sessions/<sessionId>/output`

After the final PPT is generated, the workflow moves the deliverables to
`workspace/deliverables/<sessionId>/output` and deletes the original session
workspace, including the upload cache and intermediate files. The durable
material library remains under `reference_library/`.

Typical intermediate files:

- `outline.json`
- `style.json`
- `page_notes.md`
- `document_summary.json`
- `reference_summary.json`
- `layout_options.json`
- `draft_generated.pptx`
- `rendered_preview/`

Typical archived final files:

- `workflow_generated.pptx`
- `outline.final.json`
- `style.final.json`
- `document_summary.final.json`
- `document_structure.final.json`
- `document_structure.final.md`
- `layout.final.json`
- `page_notes.final.md`
- `rendered_preview/`

## Layout Label + Diversity Validation

The brownfield layout contract keeps internal template ids stable while making
all user-facing labels Chinese:

- UI dropdowns render Chinese `displayName` / `label` text.
- `layout_options.json` keeps stable machine ids in `id` / `currentTemplate` /
  `recommendedTemplate`.
- download and archive manifests keep Chinese titles while preserving stable
  machine-readable `type` fields.

Useful local verification commands:

```bash
node --check src/ui/public/app.js
node --check src/services/layoutSelectionService.js
node --check src/services/workflowService.js
node --check src/services/deckRendererService.js
node --test test/layoutVerification.test.js
```

## Validation Notes

The current implementation has already been updated to support:

- dynamic planning for non-Sanong documents
- automatic handling of hierarchical headings like `一、1、1.1、1.1.1`
- multi-reference extraction and merge
- real page screenshots through PowerPoint rendering
- reusable layout switching by slide type

## Known Constraints

- PowerPoint-rendered previews require Microsoft PowerPoint to be available on the machine.
- The generator is optimized for Chinese bank and finance reporting style; non-financial visual styles may need a different default layout library.
- Automatic text fitting is improved, but very dense source documents may still benefit from manual layout switching or a higher page count.

## Recommended Use Pattern

- Use the UI for real work sessions and human correction.
- Use CLI for batch generation or debugging.
- Keep adding strong reference PPTs so the master material library becomes richer over time.
- When a requirement document is dense, prefer `auto` pages first, then increase to `7` to `10` if the draft still looks crowded.
