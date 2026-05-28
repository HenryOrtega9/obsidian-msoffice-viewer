# Microsoft Office Viewer (`obsidian-msoffice-viewer`)

Read-only Obsidian viewer for `.docx`, `.pptx`, `.xlsx`, and `.xls`, with zoom, an "Open in Word / PowerPoint / Excel" handoff, and an optional Claude-assisted command to **create** new Office files from a prompt. Desktop-only.

## Scope

This plugin renders Office files inside Obsidian for **reading**. It is not a WYSIWYG editor and does not modify the content of existing Office files.

### What it covers

- **Multi-format read-only preview:** `.docx`, `.pptx`, `.xlsx`, `.xls`.
- **Zoom** in / out / reset (commands + toolbar), with a configurable default zoom.
- **Open-external handoff:** a toolbar button opens the file in the native app (Word / PowerPoint / Excel).
- **Claude-assisted file creation** (optional): a command that asks Claude to generate a brand-new `.docx`, `.pptx`, or `.xlsx` from a natural-language description, then opens it in the viewer.

### Scope change from the original concept

The original concept (preserved in git history) was narrower in format but deeper in editing: a **docx-only** viewer with **Claude-mediated in-place editing** (select a span → ask Claude to rewrite it → patch the underlying `word/document.xml` or round-trip through markdown/pandoc). That editing path was **not built and is out of scope.** There is no selection capture, no OOXML run-splitting, no mammoth/pandoc round-trip, no `EditDispatcher`.

Two things changed direction instead:

1. **Wider, not deeper.** Scope broadened from docx-only to a four-format read-only viewer (docx, pptx, xlsx, xls). Editing fidelity stopped being the hard problem; faithful *rendering* across formats became the focus.
2. **Claude creates, it does not edit.** The only Claude touchpoint is the "Create Office file with Claude" command, which produces a *new* file from a structured spec. Existing files are never rewritten by Claude.

> Note on naming: internal view-type constants (`docx-claude-view`, `pptx-claude-view`, `xlsx-claude-view`) and the `docx-claude-` CSS class prefix are retained from the earlier name to avoid churn and to keep persisted workspace/view state stable. They are intentionally left as-is despite the plugin's broader scope.

---

## Architecture

### High-level flow (viewing)

```
User clicks a .docx / .pptx / .xlsx / .xls file in the vault
        │
        ▼
main.ts registers each extension → a dedicated view type
        │
        ▼
OfficeFileView (shared base: toolbar, zoom, scroll, open-external)
        │
   ┌────┴───────────────┬─────────────────────┐
   ▼                    ▼                     ▼
DocxPreviewView    PptxPreviewView      XlsxPreviewView
(docx-preview)     (native OOXML        (exceljs parse →
                    render + PDF         custom HTML grid)
                    fallback)
```

### High-level flow (Claude creation)

```
Command: "Create Office file with Claude"
        │
        ▼
CreateModal  ── user describes the file ──►  ClaudeBridge
        │                                         │
        │                          looks up app.plugins.plugins["claude-cli-chat"]
        │                          calls runHeadlessPrompt(system, user)
        │                                         ▼
        │                          Claude returns a JSON spec (fenced or raw)
        ▼                                         │
extractJsonFromResponse + validateSpec  ◄────────┘
        │   (DocxSpec | PptxSpec | XlsxSpec)
        ▼
generators/buildFile(spec) → ArrayBuffer
   docx     → docxGenerator (docx)
   pptx     → pptxGenerator (pptxgenjs)
   xlsx     → xlsxGenerator (exceljs)
        │
        ▼
write new file into the vault → open it in the matching view
```

### File layout

```
obsidian-msoffice-viewer/
  manifest.json
  package.json
  esbuild.config.mjs
  tsconfig.json
  styles.css                 # class prefix kept as docx-claude- (see naming note)
  src/
    main.ts                  # Plugin entry: registers views, extensions, commands, settings
    OfficeFileView.ts        # Shared base FileView: toolbar, zoom, scroll, open-external
    DocxPreviewView.ts       # docx-claude-view: renders via docx-preview (+ fidelity toggle)
    PptxPreviewView.ts       # pptx-claude-view: native OOXML render with PDF/fallback path
    XlsxPreviewView.ts       # xlsx-claude-view: exceljs parse → custom HTML grid
    officeToPdf.ts           # PDF rendering engine (pdf.js), used as a render/fallback path
    settings.ts              # Settings tab + defaults + zoom clamping
    ClaudeBridge.ts          # Bridge to claude-cli-chat's runHeadlessPrompt (creation only)
    CreateModal.ts           # Prompt UI for "Create Office file with Claude"
    generators/
      index.ts               # buildFile(spec) dispatch
      spec.ts                # Spec types + validation + JSON extraction from Claude output
      prompts.ts             # System prompts that instruct Claude to emit a spec
      docxGenerator.ts       # Spec → .docx (docx)
      pptxGenerator.ts       # Spec → .pptx (pptxgenjs)
      xlsxGenerator.ts       # Spec → .xlsx (exceljs)
    docx/                    # docx-preview options, feature detection, empty/warn handling
    pptx/                    # Native pptx OOXML renderer:
                             #   presentation, slide, shapes, text, tables, images,
                             #   charts/, themes, inheritance, geometry, colors, background, ooxml
    xlsx/                    # Native xlsx renderer:
                             #   grid, cells, merges, panes, richText, hyperlinks, notes,
                             #   images, themes, colors, geometry,
                             #   charts/ (chart.js), conditionalFormatting/
```

### Key dependencies

- `docx-preview`: render docx to HTML/CSS (read-only).
- `exceljs`: parse xlsx/xls workbooks for the custom grid renderer.
- `numfmt`: Excel number-format string evaluation for cell display.
- `chart.js`: render embedded charts (xlsx and pptx).
- `pdfjs-dist`: PDF rendering engine / fallback (worker source inlined at build time in `esbuild.config.mjs`).
- `jszip`: read OOXML zip archives.
- `docx`, `pptxgenjs`, `exceljs`: generate new files in the Claude creation path.
- `pptxviewjs`: pptx rendering support.

### Module responsibilities

**`main.ts`**: registers the three view types, maps extensions to them (`docx`→docx view, `pptx`→pptx view, `xlsx`/`xls`→xlsx view), each `registerExtensions` wrapped in its own try/catch so a collision with another plugin owning an extension doesn't abort the rest of load. Registers zoom commands, the "Create Office file with Claude" command, and the settings tab. On unload, detaches any open leaves of these view types.

**`OfficeFileView.ts`**: shared base for all three views: builds the toolbar (zoom buttons, zoom %, engine badge, open-external button), the scroll/render containers, and the zoom plumbing (`zoomIn` / `zoomOut` / `resetZoom`). Holds the pdf.js render helpers used by views that fall back to a PDF stage.

**`DocxPreviewView.ts`**: view type `docx-claude-view`. Reads the file binary and renders it with `docx-preview`. Exposes a fidelity toggle.

**`PptxPreviewView.ts`**: view type `pptx-claude-view`. Renders slides via the native OOXML renderer in `src/pptx/`, with a PDF stage / fallback card when native rendering can't represent a slide.

**`XlsxPreviewView.ts`**: view type `xlsx-claude-view`. Parses the workbook with `exceljs` and renders each sheet as an HTML grid (`src/xlsx/`): merges, frozen panes, rich text, hyperlinks, cell notes, images, conditional formatting, and charts. Sheet tabs allow switching sheets and toggling hidden sheets.

**`ClaudeBridge.ts`**: resolves `app.plugins.plugins["claude-cli-chat"]` at runtime and calls its exported `runHeadlessPrompt(systemPrompt, userPrompt, opts)`. Surfaces a clear `ClaudeBridgeError` if the plugin is missing, disabled, or too old to expose that method. Wraps the call in a local timeout (default 180s) so a hung subprocess can't deadlock the UI.

**`CreateModal.ts`**: prompt UI for the creation command. Sends the user's description through `ClaudeBridge`, extracts and validates the returned spec, builds the file, and hands the new `TFile` back to `main.ts` to open.

**`generators/spec.ts`**: the contract between Claude and the generators. Defines `DocxSpec` / `PptxSpec` / `XlsxSpec`, strictly validates Claude's output (`validateSpec`, throwing `SpecValidationError` on shape mismatch), and robustly extracts the JSON object from Claude's response (`extractJsonFromResponse`: strips code fences and walks for the first balanced top-level `{…}`).

**`generators/{docx,pptx,xlsx}Generator.ts`**: turn a validated spec into an Office file via `docx`, `pptxgenjs`, and `exceljs` respectively. `generators/prompts.ts` holds the system prompts that instruct Claude to emit exactly the spec shape `spec.ts` expects.

### Integration with claude-cli-chat

The creation feature depends on the `claude-cli-chat` plugin exposing `async runHeadlessPrompt(systemPrompt, userPrompt, opts?): Promise<string>`. `ClaudeBridge` is the single integration point; if claude-cli-chat is absent, disabled, or lacks that method, the command fails gracefully with an actionable error and the rest of the plugin (all viewing) is unaffected.

---

## Critical files

- `src/main.ts`: registration glue (views, extensions, commands).
- `src/OfficeFileView.ts`: shared view chrome and zoom.
- `src/XlsxPreviewView.ts` + `src/xlsx/*`: the largest renderer; most rendering complexity lives here.
- `src/pptx/*`: native pptx OOXML renderer (the focus of ongoing fidelity work).
- `src/ClaudeBridge.ts` + `src/generators/spec.ts`: the creation path's integration point and validation contract.

---

## Known risks / limitations

1. **Rendering fidelity is approximate.** docx-preview, the native pptx renderer, and the xlsx grid do not match the host app's layout 1:1. This is acceptable for reading; use the open-external button for exact layout. Document in the README.
2. **claude-cli-chat dependency (creation only).** If it's missing or lacks `runHeadlessPrompt`, the create command fails gracefully; all viewing still works.
3. **Claude spec mismatch.** Claude can emit malformed or off-shape JSON. Mitigated by `extractJsonFromResponse` (fence/brace-aware) + strict `validateSpec`, surfacing a clear error rather than writing a broken file.
4. **Desktop-only.** `isDesktopOnly: true`; the render stack (pdf.js worker, native deps) targets the desktop runtime.
5. **Extension collisions.** Another plugin may already own `.xls`/`.xlsx`/etc.; each `registerExtensions` is isolated so one collision doesn't break the others (a warning is logged).

---

## Verification

1. **Install + open.** Build (`npm run build`), copy to `<vault>/.obsidian/plugins/obsidian-msoffice-viewer/`, enable. Drop one of each format in the vault and click it: confirm each renders in its view.
2. **docx.** Open a formatted `.docx`; confirm headings, lists, tables, and the fidelity toggle render.
3. **pptx.** Open a multi-slide `.pptx`; confirm slides, text, images, tables, and charts render, and that the fallback path engages cleanly on unsupported content.
4. **xlsx / xls.** Open a multi-sheet workbook; confirm sheet tabs, merges, frozen panes, number formatting, conditional formatting, charts, notes, and hyperlinks. Confirm hidden-sheet toggle works.
5. **Zoom.** Run zoom in / out / reset commands and toolbar buttons; confirm the default-zoom setting is applied on open.
6. **Open external.** Click the open-external button; confirm the file opens in the native Office app.
7. **Create (Claude path).** Run "Create Office file with Claude" for each kind (docx/pptx/xlsx); confirm a valid file is generated, written to the vault, and opened in the matching view.
8. **Create with claude-cli-chat absent.** Disable claude-cli-chat and run the create command; confirm a clear, actionable error and that viewing is unaffected.
9. **Malformed spec.** Force a bad/fenced Claude response; confirm `validateSpec` rejects it with a readable message and no file is written.
