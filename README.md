# Obsidian Microsoft Office Viewer

An Obsidian plugin that renders Microsoft Office files (`.docx`, `.pptx`, `.xlsx`, `.xls`) read-only inside the vault, with zoom, an "Open in Word/PowerPoint/Excel" handoff, and Claude-driven creation of new Office files.

Obsidian treats Office files as opaque binaries. This plugin registers native views for them so a vault that holds coursework, financial models, and slide decks stays browsable without leaving the app.

## Features

- **Word (`.docx`)**: full document rendering via `docx-preview`, matching page layout, styles, tables, and images.
- **PowerPoint (`.pptx`)**: slide rendering backed by a local LibreOffice conversion to PDF, drawn with PDF.js for faithful output (fonts, SmartArt, charts).
- **Excel (`.xlsx`, `.xls`)**: an ExcelJS-driven grid with sheet tabs, frozen panes, merged cells, and cell formatting.
- **Zoom controls** and keyboard navigation in every view.
- **Open in native app** button for the cases where read-only is not enough.
- **Claude-driven creation**: a modal that asks Claude Code to generate a new `.docx`/`.pptx`/`.xlsx` from a prompt, written straight into the vault (design notes in `PLAN.md`).

## Architecture

- `src/main.ts`: plugin entry; registers one `OfficeFileView` per extension.
- `src/DocxPreviewView.ts`, `src/PptxPreviewView.ts`: per-format view classes.
- `src/officeToPdf.ts`: LibreOffice (`soffice`) headless conversion pipeline with caching.
- `src/docx/`, `src/pptx/`, `src/generators/`: format-specific rendering and the Claude file-creation generators.
- `src/ClaudeBridge.ts`: shells out to the Claude Code CLI for file generation.

The design deliberately prefers preview plus Claude-mediated edits over building a WYSIWYG editor; the native apps remain the editors of record.

## Requirements

- Obsidian 1.5+ (desktop only)
- LibreOffice (`soffice` on PATH) for `.pptx` rendering
- Claude Code CLI (optional, only for the file-creation modal)

## Build

```sh
npm install
npm run build
```

## License

MIT
