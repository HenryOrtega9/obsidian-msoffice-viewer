# Render Accuracy Remediation Plan

Scope: render accuracy (visual fidelity) for the three Office formats handled by this plugin, prioritized XLSX (has a visible cutoff bug) and PPTX (hard) first, DOCX last. All findings below were re-verified against the live code and, for the XLSX cutoff, against the installed `soffice` binary. File:line references are current as of this commit.

---

## 1. Executive Summary and Strategic Call

### The architecture is already the right one

The premise that the high-fidelity LibreOffice path needs to be opted into is inverted by the code. The shared `LibreOffice -> PDF -> pdf.js` pipeline (`src/officeToPdf.ts`, `src/OfficeFileView.renderViaLibreOfficePdf` at `src/OfficeFileView.ts:105-152`) is already the PRIMARY, default-first renderer for all three formats:

- `XlsxPreviewView.renderFile` calls `findSoffice()` first and renders via PDF whenever soffice exists (`src/XlsxPreviewView.ts:101-115`); the ExcelJS grid is a fallback for missing soffice and only handles `.xlsx`.
- `PptxPreviewView` tries LibreOffice, then the native OOXML DOM renderer, then pptxviewjs.
- `DocxPreviewView` defaults `mode='pdf'`, with a user toggle back to docx-preview.

soffice is present on this machine (`/opt/homebrew/bin/soffice` and `/Applications/LibreOffice.app/...`), so the PDF path is what actually renders today. The hand-rolled renderers (ExcelJS grid, native pptx DOM) are the fallbacks. This matters: most of the per-format fidelity findings below only affect users WITHOUT LibreOffice, while the cutoff bug affects EVERYONE because it lives in the PDF path.

### Strategic recommendation per format

The right model is the one DOCX already uses: PDF-first for fidelity, with an explicit per-view toggle to the interactive/selectable renderer, and the hand-rolled renderer kept (and incrementally improved) as the no-soffice fallback. Concretely:

- **XLSX**: keep `LibreOffice -> PDF` as the default. Fix the print-area cutoff in the PDF path (Section 2) using `SinglePageSheets:true` so the full used range renders instead of clipping to the print area. Add an "Interactive grid" toggle mirroring `DocxPreviewView`'s `userForcedMode` pattern so users can recover sheet tabs, cell hyperlinks, internal-link navigation, and text selection (which the PDF cannot provide). Keep the ExcelJS grid as the offline fallback and invest in its geometry/style fidelity (Section 3) since `.xls` has no other fallback and grid is what offline users see.
- **PPTX**: keep `LibreOffice -> PDF` as the default high-fidelity engine. The native DOM renderer cannot match PowerPoint on autofit metrics, preset geometry, gradients, effects, or SmartArt, and those are intractable to hand-roll faithfully. Improve the native renderer for the tractable, high-frequency gaps (autofit fontScale, bodyPr insets and anchoring, line/paragraph spacing, real bullets, flip transforms) so the no-soffice fallback is usable, and route genuinely hard slides (charts/SmartArt/3D/effects/custom geometry) to the PDF path.
- **DOCX**: keep docx-preview as the interactive/selectable default behind the toggle and `LibreOffice -> PDF` as the high-fidelity mode (already the default `mode`). Apply the small, safe docx-preview option corrections in Section 3.

### On QuickLook and embedded thumbnails (rejected as render engines)

- `qlmanage -p` emits an Apple-internal `public.html` `.qlpreview` bundle, not a portable PDF or image, and writes unreliably for pptx. It is not a viable faithful-preview engine; reimplementing Apple's QuickLook HTML host would be fragile and lower fidelity than the PDF already shipping.
- `qlmanage -t` produces a single first-page/first-slide PNG. Its only use is a darwin-only instant placeholder shown while the soffice conversion runs. Optional polish, low priority.
- Embedded OOXML thumbnails (`docProps/thumbnail.*`) are absent from real-world files (modern Office leaves "Save Thumbnail" off). Do not build on them.

Net: do NOT add QuickLook or embedded thumbnails as engines. The existing `officeToPdf.ts` pipeline is strictly superior (multi-page, cross-platform, deterministic, already half-built with caching, in-flight de-duplication, profile isolation, and font substitution).

---

## 2. The XLSX Cutoff Bug

### Root cause (file:line)

The active render path for `.xlsx` is `LibreOffice -> PDF -> pdf.js`, NOT the ExcelJS grid. `XlsxPreviewView.renderFile` (`src/XlsxPreviewView.ts:101-106`) probes `findSoffice()` first and, when present, calls `renderViaLibreOfficePdf(...)` and returns; the grid path at `src/XlsxPreviewView.ts:117-127` only runs when soffice is absent or the PDF path throws. The user's framing about `MIN_GRID_ROWS=50` / columns A..P (`src/xlsx/grid.ts:26-29`) describes the fallback grid and is irrelevant to the cutoff.

The conversion (`convertOfficeToPdf`, `src/officeToPdf.ts:212-275`) invokes the `calc_pdf_Export` filter with `PDF_FILTER_DATA` (`src/officeToPdf.ts:78-96`) that sets only `UseLosslessCompression` and `ReduceImageResolution`. There is no option to override the workbook's print area or force the full used range. LibreOffice Calc's default PDF export honors the sheet's defined print area and paginates to page size, exactly like Excel's print view. A workbook whose print area ends at row 16 exports a single Letter page containing rows 1-16 only; rows beyond the print area are dropped and the rest of the page is blank (rendered dark in Obsidian's themed stage). Finder QuickLook and Excel's normal grid view ignore the print area and show the full used range, which is why they look correct and the plugin looks cut off.

This was reproduced against the live `/opt/homebrew/bin/soffice`: a workbook with data in A1:B20 and `print_area='A1:B16'` exported via default `calc_pdf_Export` to a 1-page PDF whose text stops at row 16 (`Label 16`), with rows 17-20 absent. The same file exported with `SinglePageSheets:true` produced a 1-page PDF containing all 20 rows.

```
DEFAULT:    pages=1  labels_found=[1..16]      <- clipped to print area
SINGLEPAGE: pages=1  labels_found=[1..20]      <- full used range
```

### The exact fix

In `src/officeToPdf.ts`, branch the Calc FilterData to include `SinglePageSheets:true` (the document for this option states it "ignores each sheet's paper size, print ranges and shown/hidden status"). This both removes the print-area clip AND collapses each sheet to one un-paginated page, which also fixes the separate multi-page-overflow variant (a wide/long sheet tiling across pages where later pages may render blank in the lazy renderer).

Replace the single shared `PDF_FILTER_DATA` use in the calc branch of `pdfFilterFor` (`src/officeToPdf.ts:84-96`) with a calc-specific FilterData:

```ts
const CALC_PDF_FILTER_DATA =
  '{"SinglePageSheets":{"type":"boolean","value":"true"},' +
  '"UseLosslessCompression":{"type":"boolean","value":"true"},' +
  '"ReduceImageResolution":{"type":"boolean","value":"false"}}';
// in pdfFilterFor, calc branch:
//   return `pdf:calc_pdf_Export:${CALC_PDF_FILTER_DATA}`;
```

Then bump `CONVERSION_VERSION` (`src/officeToPdf.ts:72`, currently `"3"`) to `"4"` so cached clipped PDFs (keyed with the version at `src/officeToPdf.ts:221`) are invalidated and regenerated under the new filter.

Verification: open a workbook whose print area is smaller than its used range and confirm all rows/columns render. Re-test a wide sheet to confirm it no longer tiles into many pages (it becomes one tall/wide page per sheet).

Confidence: high. Risk: low to medium (one-page-per-sheet changes page geometry and can produce a very tall single page for huge sheets; the lazy pdf.js renderer already handles tall pages, but verify a large workbook visually).

---

## 3. Per-format Ranked Backlog

Effort scale: S (under a day), M (1-3 days), L (multi-day). Each item carries a Confidence/Risk tag matching the diagnosis severity.

### 3A. XLSX

#### X1. Print-area cutoff in the PDF path (the cutoff bug)
- Files: `src/officeToPdf.ts:78-96, 72`
- Root cause: `calc_pdf_Export` runs with default page handling, honoring the print area and paginating to page size.
- Change: add `SinglePageSheets:true` to a calc-specific FilterData; bump `CONVERSION_VERSION`. See Section 2.
- Impact: high (fixes the reported visible cutoff for every soffice user).
- Effort: S. Regression risk: medium (page geometry change for very large sheets).
- Confidence: high / Risk: medium.

#### X2. Column-width pixel formula omits Excel's 5px padding and uses crude 7px/char
- Files: `src/xlsx/grid.ts:17-18, 89-95`
- Root cause: `widthPx = Math.round(widthCh * 7)` (`grid.ts:92`). exceljs `col.width` is the stored XLSX width attribute in max-digit-width units, padding-inclusive, not the 8.43 UI character value. The crude `*7` with no padding makes the default column 59px instead of the correct 64px, and the deficit accumulates left-to-right, clipping `white-space:nowrap` text early and desyncing the absolutely-positioned image/chart overlays (`cumulativeColPx`).
- Change: replace with the ECMA-376 stored-width formula. Define `const MDW = 7` (max digit width, Calibri 11 at 96 DPI) and compute `widthPx = Math.trunc(((256 * col.width + Math.trunc(128 / MDW)) / 256) * MDW)`. This self-includes padding (verify: stored `9.140625 -> 64px`). When `col.width` is undefined, render 64px (read `ws.properties.defaultColWidth` through the same formula when present, else 64). Keep `DEFAULT_COL_WIDTH_CH = 8.43` only as a documented UI reference. See Section 4 for the formula derivation. Add a unit assertion `9.140625 -> 64`.
- Impact: high (grid fallback fidelity, overlay alignment).
- Effort: S. Regression risk: low (numeric, deterministic, testable).
- Confidence: high / Risk: low.

#### X3. Row-height handling treats height 0 as "use default" and never sets height on default rows
- Files: `src/xlsx/grid.ts:110-119`
- Root cause: `const heightPx = row.height ? round(...) : DEFAULT` (`grid.ts:113-115`) collapses a legitimate 0-height (hidden/collapsed) row into the 20px default, and the `tr` height is only set when `row.height` is truthy (`grid.ts:116-118`), so padded/blank rows render at CSS content height while the cumulative math assumes 20px, drifting the overlay layer. Custom sheet default row height (`ws.properties.defaultRowHeight`) is never read.
- Change: derive `DEFAULT_ROW_HEIGHT_PX = Math.round(15 * POINTS_TO_PX)` (= 20, but documented from 15pt) and seed it from `ws.properties.defaultRowHeight` (points) when present. Replace the truthiness check: `if (row.hidden || row.height === 0) heightPx = 0; else if (row.height == null) heightPx = defaultRowPx; else heightPx = Math.round(row.height * POINTS_TO_PX);`. Always set the `tr` height inline so DOM rows match `cumulativeRowPx`.
- Impact: medium (correct row geometry, overlay alignment, hidden rows).
- Effort: S. Regression risk: low.
- Confidence: high / Risk: low.

Note: `POINTS_TO_PX = 4/3` (`grid.ts:19`) is verified correct at 96 DPI. Keep it.

#### X4. Number-format-driven font color (`[Red]`, conditional sections) never applied
- Files: `src/xlsx/cells.ts:59-96, 161-173`
- Root cause: `cellText` calls numfmt `format()` for the text but never `formatColor()`, so a cell formatted `#,##0;[Red]-#,##0` shows a negative number in the static font color instead of red. Conditional-section colors (`[<50][Red]...`) are likewise ignored.
- Change: call `numfmt.formatColor(fmt, raw, {indexColors:true, throws:false})` alongside `format()`; when non-null, set `td.style.color` to that value (overriding `font.color`). numfmt is already a dependency.
- Impact: medium to high (financial/accounting sheets, grid path).
- Effort: S. Regression risk: low.
- Confidence: high / Risk: low.

#### X5. Cheap font bugs: bold weight, strike, super/subscript, underline variants
- Files: `src/xlsx/cells.ts:161-173`, `src/xlsx/richText.ts`
- Root cause: `font.bold` maps to `font-weight: 600` (`cells.ts:166`); Excel bold is 700. `font.strike`, `font.vertAlign` (super/subscript), and underline variants (`double`/`singleAccounting`) are not handled. The same gaps exist in the rich-text run renderer.
- Change: `bold -> font-weight: bold`; `font.strike -> line-through` (merge with underline into a single `text-decoration`); `font.vertAlign === 'superscript'/'subscript' -> <sup>/<sub>` or `vertical-align`; underline `'double' -> text-decoration: underline double`. Fix in both `cells.ts` and `richText.ts`.
- Impact: medium.
- Effort: S. Regression risk: low.
- Confidence: high / Risk: low.

#### X6. Border style fidelity is coarse; default border color is wrong
- Files: `src/xlsx/cells.ts:215-221`
- Root cause: `borderCss` maps everything to `solid` except `double`, forces 1-2px, and defaults the unspecified color to `#d4d4d4`. Excel's automatic border color is black, and hair/dotted/dashed/medium/thick all render as plain thin solid lines.
- Change: a proper style table: `hair/thin -> 1px solid`, `dotted -> 1px dotted`, `dashed/dashDot/dashDotDot -> 1px dashed`, `medium/mediumDashed -> 2px solid|dashed`, `thick -> 3px solid`, `double -> 3px double`; default unspecified color to `#000000`. A shared-edge ownership pre-pass (priority ladder, higher wins, ties to the lower/right cell) avoids double-painting under `border-collapse:collapse`. The color table is a safe numeric change; the ownership pass is layout-affecting (verify visually).
- Impact: low to medium (ruled financial tables).
- Effort: S (style table) + M (ownership pass). Regression risk: low (table) to medium (ownership).
- Confidence: high / Risk: low.

#### X7. Alignment: indent, text rotation, vertical/stacked text, shrink-to-fit ignored
- Files: `src/xlsx/cells.ts:194-206`
- Root cause: `cellInlineStyle` maps only horizontal/vertical/wrapText. `alignment.indent`, `textRotation`, `shrinkToFit`, and `fill`/`centerContinuous`/`distributed` modes are dropped; indented label columns and rotated headers (common in financial models) render flat and left-flush.
- Change: `indent -> padding-left/right of indent*~9px`; `textRotation 1..90 -> rotate(-deg)`, `91..180 -> rotate(value-90)`, `255 -> stacked vertical text` (writing-mode); `shrinkToFit -> canvas measureText + transform: scale(min(1, avail/textWidth))` with `transform-origin: left`, mutually exclusive with wrapText. Guard transforms so they do not break cell box sizing.
- Impact: medium.
- Effort: M. Regression risk: medium (transforms interact with cell box layout; verify visually).
- Confidence: medium / Risk: medium.

#### X8. Accounting/currency alignment (symbol left, digits right, trailing slot)
- Files: `src/xlsx/cells.ts`
- Root cause: accounting formats are rendered with plain `text-align:right`, which does not left-align the currency symbol to the cell edge the way Excel does.
- Change: detect accounting/currency via `numfmt.getFormatInfo().type`; render a 3-span flex cell (symbol left, digits right, trailing parenthesis slot) with `font-variant-numeric: tabular-nums`; pass numfmt `skipChar:' '` so `'_)'` reserves a parenthesis-width slot.
- Impact: medium (financial sheets).
- Effort: M. Regression risk: medium (layout; verify visually).
- Confidence: medium / Risk: medium.

#### X9. Pattern fills (non-solid) painted as flat fgColor; gradient fills dropped
- Files: `src/xlsx/cells.ts:175-184`
- Root cause: for `fill.type==='pattern'` the code paints `fgColor` solid regardless of `patternType` (gray125 becomes a saturated block); `fill.type==='gradient'` is dropped.
- Change: for non-solid patterns blend fg over bg by density (or a CSS repeating-linear-gradient); for gradient fills emit a CSS linear/radial gradient from `fill.gradient.stops`.
- Impact: low.
- Effort: M. Regression risk: low.
- Confidence: medium / Risk: low.

#### X10. Diagonal borders not rendered
- Files: `src/xlsx/cells.ts`
- Root cause: `cell.border.diagonal {up,down,style,color}` is never read.
- Change: on a `position:relative` td, inject an absolutely-positioned inline `<svg><line>` from corner to corner (handles non-square cells without trig).
- Impact: low.
- Effort: S. Regression risk: low.
- Confidence: medium / Risk: low.

#### X11. Used-range detection should consult `ws.dimensions`
- Files: `src/xlsx/grid.ts:61-70`
- Root cause: `lastRow`/`lastCol` derive from `actualRowCount`/`rowCount` plus merges. `actualRowCount` is a count of non-empty rows, not the last index; a trailing format-only/merge-only row can under-report.
- Change: also take `ws.dimensions.bottom`/`.right` into the `Math.max` (verified reliable: `dimensions={top:1,left:1,bottom:20,right:2}` for a 20-row sheet).
- Impact: low.
- Effort: S. Regression risk: low.
- Confidence: medium / Risk: low.

#### X12. General number format and `@` text format diverge from Excel
- Files: `src/xlsx/cells.ts:59-96, 152-159`
- Root cause: `formatGeneralNumber` rounds to 1e10 and stringifies (no scientific-notation switchover, no 11-sig-digit rule); `fmt==='@'` routes numbers through it; dates with General fall back to `toLocaleDateString()` (locale-dependent). numfmt does NOT support the 1904 date system.
- Change: route General numbers through numfmt's `'General'` formatter; for `@` coerce the number to its General string and treat as text; for dates with no/General fmt use a fixed Excel-like default via numfmt. When `workbook.properties.date1904` is set, add 1462 to the serial before `format()`.
- Impact: medium.
- Effort: M. Regression risk: medium (touches the value pipeline; verify a sample of number/date/text cells).
- Confidence: medium / Risk: medium.

#### X13. Auto-fit (bestFit) columns and frozen-pane offset timing
- Files: `src/xlsx/grid.ts:89-96`, `src/xlsx/panes.ts:18-57`
- Root cause: bestFit columns render at default width; frozen-pane sticky offsets read `thead.offsetHeight` synchronously before layout settles, and px offsets are unscaled under CSS zoom.
- Change: optionally measure content width via canvas for unspecified columns (capped); defer pane application to `requestAnimationFrame` and compute offsets from `getBoundingClientRect`.
- Impact: low.
- Effort: M. Regression risk: medium (layout; verify visually).
- Confidence: low / Risk: medium.

#### X14. (UX, not accuracy) add the PDF <-> grid toggle and engine label
- Files: `src/XlsxPreviewView.ts`, `src/OfficeFileView.ts:85-89`
- Change: mirror `DocxPreviewView`'s `userForcedMode` toggle so users can switch between high-fidelity PDF (default) and the interactive grid (sheet tabs, hyperlinks, selection); add `setEngineLabel('LibreOffice'/'Grid')` for parity with pptx.
- Impact: UX leverage, not pixel accuracy.
- Effort: M. Regression risk: low.
- Confidence: high / Risk: low.

### 3B. PPTX

These affect the NATIVE DOM renderer (the no-soffice fallback). The default soffice PDF path already handles all of them. Prioritize the ones with the highest visual ROI for offline users.

#### P1. normAutofit fontScale / lnSpcReduction not applied (the #1 native-renderer fidelity bug)
- Files: `src/pptx/text.ts:35-134`, `src/pptx/slide.ts:123-132`
- Root cause: PowerPoint pre-computes shrink-to-fit and writes `a:normAutofit` `fontScale`/`lnSpcReduction` into the file; the renderer never reads `a:bodyPr`/`a:normAutofit`, so titles/bodies that PowerPoint shrank render full-size and clip (the shape is `overflow:hidden`, `styles.css:221-226`).
- Change: in `renderSp` read `a:bodyPr -> a:normAutofit`; thread `{fontScale, lnSpcReduction}` into `renderParagraphsInto`. Compute `span.style.fontSize = (pt * fontScale/100000) * EMU_PER_PT * scale` and reduce line-height by `(1 - lnSpcReduction/100000)`. Values are thousandths of a percent (divide by 100000; `fontScale=62500` means 62.5%).
- Impact: high.
- Effort: M. Regression risk: medium (affects every text shape; verify visually).
- Confidence: high / Risk: medium.

#### P2. No `a:bodyPr` insets or vertical anchoring
- Files: `src/pptx/slide.ts:123-132`, `src/pptx/text.ts`
- Root cause: `renderSp` never reads `a:bodyPr`, so there are no text insets (default `lIns=91440`, `tIns=45720`, `rIns=91440`, `bIns=45720` EMU) and no vertical anchor; text hugs the shape edge and master title placeholders (usually `anchor=ctr`) render top-aligned.
- Change: apply insets as CSS padding (`lIns * scale`px etc.); make the shape div `display:flex; flex-direction:column` and map `anchor` t/ctr/b to `justify-content` flex-start/center/flex-end; honor `anchorCtr=1` with `align-items:center`. Resolve insets/anchor through the placeholder inheritance chain (slide, then layout, then master, then defaults).
- Impact: critical (native path).
- Effort: M. Regression risk: medium (layout restructure; verify visually).
- Confidence: high / Risk: medium.

#### P3. No line spacing or paragraph spacing
- Files: `src/pptx/text.ts:35-54, 108-134`
- Root cause: `a:lnSpc`/`a:spcBef`/`a:spcAft` are never read; line-height is hardcoded 1.2 (`styles.css:225`) and paragraph margin 0.
- Change: parse `a:pPr` children. `a:lnSpc` `a:spcPct val/100000 -> unitless line-height`; `a:spcPts val/100 *EMU_PER_PT*scale -> px line-height`. `spcBef -> margin-top`, `spcAft -> margin-bottom` from `spcPts/100` (pt) or `spcPct/100000` of font size. If both autofit lnSpcReduction and percent lnSpc are present, multiply.
- Impact: high.
- Effort: M. Regression risk: medium (vertical rhythm; verify visually).
- Confidence: high / Risk: medium.

#### P4. No real bullets, auto-numbering, or hanging indents
- Files: `src/pptx/text.ts:108-112`, `src/pptx/inheritance.ts`
- Root cause: `a:buChar`/`a:buAutoNum`/`a:buNone`/`marL`/`indent` are ignored; the only list cue is `marginInlineStart = level*1.2em` (`text.ts:112`), and no bullet glyph is emitted.
- Change: parse `marL`/`indent` (EMU) -> `paddingLeft = marL*scale`, `textIndent = indent*scale` (negative for hanging); parse `buChar`/`buAutoNum`(type)/`buNone`/`buFont`/`buSzPct(/100000)`/`buSzPts(/100)`/`buClr` and render the bullet as a positioned span at the hang position; maintain a per-level counter for `buAutoNum`. Inherit bullet defaults from the master/layout body `lstStyle` `lvlNpPr`. Drop the `level*1.2em` hack.
- Impact: high (bulleted content lists are the most common slide body).
- Effort: M to L. Regression risk: medium (layout; verify visually).
- Confidence: high / Risk: medium.

#### P5. Layout/master background shapes never drawn
- Files: `src/pptx/slide.ts:65-66`
- Root cause: `renderSlide` walks only the slide's own spTree, so master/layout logos, footers, date/slide-number placeholders, and decorative shapes are absent; branded templates render blank.
- Change: composite master spTree, then layout spTree, then slide spTree (honor `showMasterSp`, default 1). Skip placeholder shapes re-supplied by a more specific level (match by type+idx), but draw non-placeholder graphics. Gate behind z-order (master lowest).
- Impact: high.
- Effort: L. Regression risk: high (compositing change, double-draw hazards; verify visually).
- Confidence: high / Risk: high.

#### P6. Preset geometry ignored: every shape renders as a rectangle
- Files: `src/pptx/slide.ts:108-121`, `src/pptx/shapes.ts`
- Root cause: `a:prstGeom`/`a:custGeom` are never read; ellipses, rounded rects, triangles, arrows, callouts, stars all render as sharp rectangles.
- Change: handle high-frequency presets with CSS (`ellipse -> border-radius:50%`, `roundRect -> border-radius from adj`) and clip-path/SVG paths for triangles/arrows/etc.; build a small RPN guide-formula evaluator for the common ~20-40 shapes. Unmapped presets and `custGeom` fall back to the rectangle (acceptable) or route to PDF.
- Impact: high (diagram-heavy slides).
- Effort: L. Regression risk: medium (verify visually).
- Confidence: high / Risk: medium.

#### P7. tint/shade applied in raw sRGB instead of linear scRGB
- Files: `src/pptx/colors.ts:122-145`
- Root cause: `applyMods` does `shade: r*=shade` and `tint: r=r*tint+255*(1-tint)` in 0..255 sRGB (`colors.ts:124-133`). DrawingML `a:tint`/`a:shade` operate in LINEAR scRGB; the sRGB approximation produces visibly wrong midtones. (Distinct from the SpreadsheetML HSL-luminance tint in `src/xlsx/colors.ts:69-117`, which is correct for xlsx and must NOT be shared.)
- Change: add sRGB<->scRGB transfer functions (`c<=0.04045 ? c/12.92 : ((c+0.055)/1.055)^2.4` and inverse), then `applyShade: scRGB*=shade`, `applyTint: scRGB=1-(1-scRGB)*tint` between linearize and delinearize, matching Apache POI `DrawPaint`. Keep `lumMod`/`lumOff`/`sat`/`hue` in HSL after delinearize.
- Impact: medium.
- Effort: M. Regression risk: medium (color correctness; verify against a known theme swatch).
- Confidence: high / Risk: medium.

#### P8. satMod/satOff/hueMod/hueOff not implemented
- Files: `src/pptx/colors.ts:73-88, 122-145`
- Root cause: `readMods`/`applyMods` handle only tint/shade/lumMod/lumOff/alpha; chart and accent-derived fills use sat/hue mods and render with wrong saturation.
- Change: read `satMod`/`satOff`/`hueMod`/`hueOff`; multiply sat by `satMod/100000` and add `satOff` (same for hue) in HSL, after tint/shade.
- Impact: medium.
- Effort: S. Regression risk: low (verify the lumOff scale matches the 0..1 HSL representation; see Section 4).
- Confidence: medium / Risk: low.

#### P9. Gradient and picture/pattern fills collapse to one color or white
- Files: `src/pptx/shapes.ts:8-18`, `src/pptx/background.ts:30-53`
- Root cause: `resolveShapeFill` returns a gradient's first stop only (`shapes.ts:13-16`); blipFill on a shape is unhandled; backgrounds return null for image/gradient.
- Change: build a CSS linear/radial gradient from all `a:gs` stops (`pos/1000 = %`), mapping `a:lin@ang` (60000ths deg) to `cssAngle=(ang/60000+90) mod 360`; for `a:blipFill` on a shape reuse the image-loading path with `background-size`; resolve image backgrounds to an object URL. Map `a:pattFill` pctNN/diagonal/grid presets to SVG `<pattern>` tiles (today they render blank).
- Impact: medium.
- Effort: M. Regression risk: low.
- Confidence: high / Risk: low.

#### P10. flipH/flipV not applied to shapes and pictures
- Files: `src/pptx/slide.ts:227-233, 135-140`
- Root cause: `applyRotation` reads only `rot`; `renderPic` reads neither flip nor rotation. Mirrored arrows and flipped/rotated images render wrong.
- Change: compose `transform: rotate(deg) scaleX(flipH?-1:1) scaleY(flipV?-1:1)` with `transform-origin: center`; apply the same to the picture `<img>`. Propagate group `rot`/`flip` into the child frame.
- Impact: medium.
- Effort: S to M. Regression risk: low.
- Confidence: high / Risk: low.

#### P11. Image cropping (`a:srcRect`), rotation, and aspect ratio wrong
- Files: `src/pptx/images.ts:28-68`, `src/pptx/slide.ts:135-140`, `styles.css:238-243`
- Root cause: `srcRect` (l/t/r/b crop in thousandths) is never read, so cropped images show full extent; `object-fit:fill` distorts aspect ratio; pic `xfrm` rot/flip ignored.
- Change: when `srcRect` present, wrap the img in `overflow:hidden` and translate/scale so the cropped region fills the box (or `object-fit:cover` with `object-position` from the crop); keep `object-fit:fill` only when no crop. Apply rot/flip from the pic xfrm.
- Impact: medium.
- Effort: M. Regression risk: low.
- Confidence: high / Risk: low.

#### P12. Font resolution lacks substitution and ea/cs scripts
- Files: `src/pptx/inheritance.ts:115-123`, `src/pptx/text.ts:84-88`
- Root cause: `resolveThemeFont` emits `"Name", sans-serif`; uninstalled Calibri/Cambria/Aptos silently fall to sans-serif (different metrics, different line breaks). The soffice path substitutes Calibri->Carlito etc. (`officeToPdf.ts:103-108`) but the native renderer does not. `a:ea`/`a:cs` scripts are dropped.
- Change: mirror the `FONT_SUBSTITUTIONS` table (Calibri->Carlito, Cambria->Caladea, "Aptos Display"->Aptos) and append metric-compatible fallbacks; seed a default font on the stage from `theme.minorFont`; resolve `+mj-ea/+mn-ea/+mj-cs/+mn-cs` and run-level `a:ea/a:cs`. Consider bundling Carlito/Caladea as `@font-face` for deterministic metrics.
- Impact: medium.
- Effort: M. Regression risk: low.
- Confidence: high / Risk: low.

#### P13. Placeholder/style inheritance is partial
- Files: `src/pptx/inheritance.ts:65-111`, `src/pptx/text.ts:56-91`
- Root cause: `buildPhDefaults` reads only the master `p:txStyles` (`inheritance.ts:71, 95-106`), skipping italic/underline, the layout/master placeholder `lstStyle`, the slide shape `a:lstStyle`, and `presentation.xml` `defaultTextStyle`. `findPhShape` matches type+idx then type-only; per spec idx should be preferred. `pickStyleEl` routes only title/ctrTitle to titleStyle (correct) but subTitle/obj should map to bodyStyle.
- Change: extend `RunStyleDefaults` with italic/underline; layer inheritance master txStyles -> layout/master placeholder lstStyle -> slide shape lstStyle -> run rPr, each filling only unset properties; prefer idx in `findPhShape`; route subTitle/obj to bodyStyle.
- Impact: medium.
- Effort: M to L. Regression risk: medium (resolution order changes visible styling; verify visually).
- Confidence: medium / Risk: medium.

#### P14. Text overflow clipped with no shrink-to-fit safety net
- Files: `styles.css:221-226`, `src/pptx/text.ts`
- Root cause: shape is `overflow:hidden` with no autofit; combined with P1-P3 metric differences, content that fits in PowerPoint can overflow and vanish.
- Change: implement P1 (stored fontScale). As a net, after layout, if `scrollHeight>clientHeight` on an autofit body, iteratively reduce font-size until it fits. Keep `overflow:hidden` for `noAutofit` shapes.
- Impact: medium.
- Effort: M. Regression risk: medium (verify visually).
- Confidence: high / Risk: medium.

#### P15. Default slide size fallback is 4:3
- Files: `src/pptx/presentation.ts:31-51`, `src/pptx/render.ts`
- Root cause: when `p:sldSz` is absent (rare but legal) the default is hardcoded 4:3 (9144000x6858000); modern decks are 16:9.
- Change: default to 16:9 (cx=12192000) when `sldSz` is absent. The fit-to-width scale itself is correct; no other change needed.
- Impact: low.
- Effort: S. Regression risk: low.
- Confidence: medium / Risk: low.

#### P16. Table fidelity gaps (tableStyle, cell margins, default font)
- Files: `src/pptx/tables.ts:11-102`
- Root cause: only explicit per-cell fills/borders are honored; `a:tblPr` `firstRow`/`bandRow`/`firstCol` and the referenced `tableStyleId -> tableStyles.xml` (header fill, banding, theme borders) are ignored; cell `marL/marR/marT/marB` are ignored (fixed CSS padding); cell text gets no `defaultsFor` so it falls to 18pt.
- Change: resolve `tblPr` style id against `ppt/tableStyles.xml` and apply band/header/whole-table fills/borders; read cell margins as padding; pass a body-style `defaultsFor` into cell `parseTxBody`.
- Impact: medium.
- Effort: L. Regression risk: medium (verify visually).
- Confidence: high / Risk: medium.

#### P17. Shape line dash/cap/join ignored
- Files: `src/pptx/shapes.ts:25-34`, `src/pptx/slide.ts:116-120`
- Root cause: `resolveShapeLine` reads only solidFill color and width; `a:prstDash`, gradient line fills, cap/cmpd are dropped and the border is hardcoded `solid`.
- Change: read `a:prstDash` and map to CSS `border-style` (dash->dashed, dot->dotted). For non-rectangular shapes move the stroke to the SVG/clip-path (depends on P6).
- Impact: low.
- Effort: S. Regression risk: low.
- Confidence: medium / Risk: low.

#### P18. Charts via Chart.js are approximate (accept and document)
- Files: `src/pptx/charts/render.ts:14-46`
- Root cause: charts render through the xlsx Chart.js path; Office chart styling (data labels, axis styling, 3D/combo/radar) cannot be matched.
- Change: accept as best-effort for the native tier and document it; rely on the soffice PDF path for accurate charts. Optionally forward explicit per-series `c:spPr` fills.
- Impact: low.
- Effort: S (doc). Regression risk: low.
- Confidence: medium / Risk: low.

### 3C. DOCX (lowest priority)

docx-preview is the HTML fallback behind the toggle; the PDF path is the default. All findings are option-level or routing.

#### D1. `renderComments`/`renderChanges` enabled (shows review markup as content)
- Files: `src/docx/options.ts:23-24`
- Root cause: both are set true, overriding docx-preview defaults (both false). `renderChanges:true` renders tracked insertions AND deletions inline, which does not match Word's default final view; `renderComments` relies on the CSS Highlight API and does not reproduce margin bubbles.
- Change: for a read-only viewer mirroring the final document, set `renderChanges:false` (clean final text, matching Word/LibreOffice default). Set `renderComments:false` unless comment visibility is a deliberate feature.
- Impact: medium-attention (changes visible output; confirm desired behavior).
- Effort: S. Regression risk: low.
- Confidence: high / Risk: low.

#### D2. `hideWrapperOnPrint` not set
- Files: `src/docx/options.ts:7-25`
- Root cause: defaults to false; with `inWrapper:true` the page-shadow chrome is kept when printing/exporting from the HTML fallback.
- Change: add `hideWrapperOnPrint:true`. Print-path only; no on-screen effect.
- Impact: low.
- Effort: S. Regression risk: low.
- Confidence: medium / Risk: low.

#### D3. `experimental:true` and `useBase64URL:true` are correct; keep them
- Files: `src/docx/options.ts:15, 18`
- These match the recommended read-only-fidelity set (tab-stop calculation; base64 data URLs survive Obsidian/Electron CSP and avoid object-URL revocation). No change.

#### D4. `featureDetect.ts` is dead code (complex-feature PDF routing never wired in)
- Files: `src/docx/featureDetect.ts:20`, `src/DocxPreviewView.ts:46-73`
- Root cause: `detectComplexFeatures()`/`DocxComplexity` (with `forcePdf` for charts/SmartArt/OLE) are defined but never imported; `renderFile` routes purely on `userForcedMode ?? "pdf"`. docx-preview cannot render charts/SmartArt/OLE, so toggling to HTML on such a doc silently omits them.
- Change: wire the detector into the HTML branch, short-circuiting to PDF (or surfacing a one-line notice) when `forcePdf` and soffice is available; OR delete `featureDetect.ts` if the routing is intentionally abandoned. Behavioral routing decision, confirm intent first.
- Impact: medium (only the manually-toggled HTML path).
- Effort: S. Regression risk: low.
- Confidence: high / Risk: low.

#### D5. Single-tall-page collapse for documents lacking pagination hints
- Files: `src/docx/options.ts:14`, `src/DocxPreviewView.ts:77-93`
- Root cause: docx-preview cannot measure content height; docs lacking `lastRenderedPageBreak` (Google Docs/pandoc/python-docx) collapse into one tall page in the HTML fallback.
- Change: no safe option flip fixes this in 0.3.7. Do NOT flip `ignoreLastRenderedPageBreak` to true (it removes the only pagination signal). The PDF-first default and the toggle already mitigate; consider a subtle UX hint. Leave the option as-is.
- Impact: low. Effort: S (UX hint, optional). Confidence: medium / Risk: low.

#### D6. (perf/correctness, shared with all formats) cache key, timeout, eviction
- Files: `src/officeToPdf.ts`
- The core best practices are already present (SHA-256 content hash folded with `CONVERSION_VERSION` and extension into the key at `:219-223`, dedicated `-env:UserInstallation` profile at `:249`, in-flight de-duplication at `:210, 234-235, 269-274`). Additions: fold the LibreOffice version into the cache key so an LO upgrade invalidates stale PDFs; lower the 120s execFile timeout (`:256`) toward 30-60s for the interactive open path and quarantine repeat-failing hashes; add LRU/size-cap eviction to the flat `os.tmpdir()` cache dir (`:203-205`), optionally sharded into hash-prefix subdirs.
- Impact: low (robustness). Effort: M. Confidence: high / Risk: low.

---

## 4. Authoritative Reference Appendix

Constants and formulas, verified against the cited sources, for future work. Hyphens in code/identifiers are intentional.

### 4.1 XLSX column width (chars <-> pixels), ECMA-376

- Maximum Digit Width (MDW): the pixel width of the widest digit 0-9 in the workbook's normal font at 96 DPI. For Calibri 11, MDW = 7px. General: `MDW = ROUND((advanceWidthFUnits / unitsPerEmFUnits) * ROUND(fontSizePt/72*DpiX))`.
  Source: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
- Cell padding (PP): Excel adds a fixed 5px per column (4px margin + 1px gridline). General: `PP = 2 * CEILING(MDW/4) + 1` (= 5 for MDW 7).
  Source: https://github.com/closedxml/closedxml/wiki/Cell-Dimensions
- Stored-width to pixels (use this for exceljs `col.width`, which is the stored attribute in MDW units, padding-inclusive):
  `px = Trunc( ( ( 256 * width + Trunc(128 / MDW) ) / 256 ) * MDW )`. With MDW=7, `Trunc(128/7)=18`. Worked: stored `9.140625 -> Trunc((2358/256)*7) = Trunc(64.4766) = 64px`. The `Trunc(128/MDW)` term self-includes padding (do not also add +5).
  Source: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
- Linear approximation from a true character count: `px = round(chars * MDW + 5)` (keep +5 separate from the per-char factor).
  Source: https://docs.sheetjs.com/docs/csf/features/colprops/
- Empirical PhpSpreadsheet ratio (within 1px, embeds padding): `px = round(col.width * 64/9.140625)` (= `*7.0017` for Calibri 11). Default column = 64px (not the 59px that `round(8.43*7)` gives).
  Source: https://github.com/PHPOffice/PhpSpreadsheet/blob/master/src/PhpSpreadsheet/Shared/Drawing.php
- Rounding modes matter: the spec uses Trunc (toward zero, `Math.trunc`) for the stored-width and px formulas, but `+0.5` then Trunc for the char-count display formula. Using `Math.round` everywhere introduces off-by-one drift.
- Narrow-column branch: when target px < `MDW + PP` (= 12px) Excel switches to `chars = px / (MDW + PP)`. Edge case for very thin custom columns.
- MDW=7 is Calibri-11/96-DPI specific. For a non-Calibri normal font, measure MDW via an offscreen canvas (`measureText('0')`) and recompute `PP = 2*ceil(MDW/4)+1`.

### 4.2 Row height and units

- Points to pixels at 96 DPI: `px = round(points * 96/72) = round(points * 4/3)`. `POINTS_TO_PX = 4/3` in `grid.ts` is correct.
- Default row height (Calibri 11) = 15pt = `round(15 * 4/3) = 20px`. Derive `DEFAULT_ROW_HEIGHT_PX = round(15 * POINTS_TO_PX)`; honor `ws.properties.defaultRowHeight` when present.
- `row.height === 0` or `row.hidden` => 0px (collapsed), not the default. Only fall back to the default when `row.height` is null/undefined.
  Source: https://github.com/closedxml/closedxml/wiki/Cell-Dimensions
- Unit identities: 1pt = 1/72 in; 96px = 1in; 1pt = 12700 EMU.
  Source: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/

### 4.3 XLSX (SpreadsheetML) tint, theme order, number formats

- SpreadsheetML tint (the variant the xlsx renderer must use, in HSL luminance, 0..1 normalized): `if tint<0: L' = L*(1+tint); if tint>0: L' = L*(1-tint)+tint`. The current `src/xlsx/colors.ts:89` (`newL = tint<0 ? l*(1+tint) : l*(1-tint)+tint`) is algebraically identical to ISO/IEC 29500 with HLSMAX normalized to 1.0. KEEP IT. Do NOT switch to the per-channel `c'=c*(1-t)+t` form (that shifts hue and is non-spec).
  Source: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.color
- Theme index order: cell `theme=` indices use Excel order dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink, which differs from the `clrScheme` file order (lt1,dk1,lt2,dk2,...). The mapping in `src/xlsx/themes.ts` (lt1->1, dk1->0, lt2->3, dk2->2) is correct.
- numfmt: `formatColor(pattern, value, {indexColors:true})` returns the section's CSS color (`[Red]`, `[Color N]`, conditional sections). `getFormatInfo(pattern)` classifies type (currency/date/percent/...). numfmt does NOT support the 1904 date system: add 1462 to the serial when `date1904` is set. The `'*'` fill and `'_'` skip operators emit nothing unless `fillChar`/`skipChar` are passed.
  Source: https://github.com/borgar/numfmt/blob/master/API.md

### 4.4 XLSX border style -> CSS map

`hair->1px solid`, `thin->1px solid`, `dotted->1px dotted`, `dashed->1px dashed`, `dashDot/dashDotDot->1px dashed`, `medium->2px solid`, `mediumDashed->2px dashed`, `thick->3px solid`, `double->3px double`. Default unspecified color = `#000000` (not `#d4d4d4`). Shared-edge precedence ladder (higher wins, ties to lower/right cell): none < hair < dotted < dashDotDot < dashDot < dashed < thin < mediumDashDotDot < mediumDashDot < mediumDashed < medium < double < thick.
Source: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.color

### 4.5 PPTX (DrawingML) EMU and percent units

- `914400 EMU = 1 inch`; `12700 EMU = 1 pt`; `9525 EMU = 1 px` at 96 DPI. `EMU_PER_PX=9525`, `EMU_PER_PT=12700` in `src/pptx/geometry.ts` are correct.
  Source: https://web.mit.edu/~stevenj/www/ECMA-376-new-merged.pdf
- bodyPr inset defaults: `lIns=91440` (0.1in), `rIns=91440`, `tIns=45720` (0.05in), `bIns=45720` EMU.
  Source: https://www.datypic.com/sc/ooxml/e-a_bodyPr-1.html
- Vertical anchor `a:bodyPr@anchor`: t/ctr/b/just/dist; default t. Map to flex `justify-content`.
  Source: https://www.datypic.com/sc/ooxml/t-a_ST_TextAnchoringType.html
- normAutofit: `fontScale` and `lnSpcReduction` are thousandths of a percent (divide by 100000). `fontScale` default 100000; `lnSpcReduction` default 0. `effectivePt = pt * fontScale/100000`; `lineHeight *= 1 - lnSpcReduction/100000`.
  Source: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextSpacingPercen_topic_ID0EO4TOB.html
- Line spacing `a:lnSpc`: `a:spcPct val/100000 -> unitless line-height`; `a:spcPts val/100 (pts) -> *EMU_PER_PT*scale px`. Paragraph spacing `a:spcBef`/`a:spcAft`: same spcPts/spcPct rules -> margin-top/bottom.
  Sources: https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_lnSpc_topic_ID0E3KTKB.html , https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_spcPct_topic_ID0EZ4WKB.html
- Bullets: `marL`/`indent` in EMU (`paddingLeft = marL*scale`, `textIndent = indent*scale`, negative for hanging). `buAutoNum` types: arabicPeriod, arabicParenR, alphaLcParenR, alphaUcPeriod, alphaLcPeriod, romanLcPeriod, romanUcPeriod. `buSzPct/100000`, `buSzPts/100`. Typical master lvl1: `marL=457200 indent=-457200` (0.5in hang).
  Source: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_buAutoNum_topic_ID0EQZALB.html
- Rotation `a:xfrm@rot` in 60000ths of a degree (`degFromRot = rot/60000`, correct). Rotation is about the box center (`transform-origin: center`). Compose flips: `transform: rotate(deg) scaleX(flipH?-1:1) scaleY(flipV?-1:1)`.
  Source: https://www.datypic.com/sc/ooxml/e-a_xfrm-1.html
- Default slide size when `p:sldSz` absent: 4:3 = `9144000 x 6858000` EMU. Widescreen 16:9 = `12192000 x 6858000` (modern default).
  Source: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/de4053e1-e1ca-4cc2-a810-aa2da4cc43a0

### 4.6 PPTX (DrawingML) color transforms

- Color resolution order: `schemeClr val -> clrMap remap (only bg1/tx1/bg2/tx2) -> theme clrScheme slot -> modifiers`. accent1-6/hlink/folHlink map 1:1. `clrMap` lives on the slideMaster, overridable per-layout/slide via `clrMapOvr`. `phClr` resolves from the `<a:style>` fillRef/lnRef/effectRef + fmtScheme, not the theme.
  Source: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_clrMap_topic_ID0ETDFMB.html
- tint/shade in LINEAR scRGB (Apache POI DrawPaint reference): (1) linearize each sRGB channel `c<=0.04045 ? c/12.92 : ((c+0.055)/1.055)^2.4`; (2) `applyShade: scRGB *= shade/100000`; (3) `applyTint: scRGB = 1 - (1 - scRGB) * (tint/100000)`; (4) delinearize `c<=0.0031308 ? c*12.92 : 1.055*c^(1/2.4) - 0.055`; (5) HSL mods; (6) HSL->RGB.
  Source: https://raw.githubusercontent.com/apache/poi/trunk/poi/src/main/java/org/apache/poi/sl/draw/DrawPaint.java
- HSL mods order hue, sat, lum: `if mod != -1: hsl[part] *= mod/100000; if off != -1: hsl[part] += off`. On a 0..100 HSL scale POI divides the offset by 1000 (lumOff=20000 -> +20 L points). On a 0..1 HSL scale (as `src/pptx/colors.ts` uses) the offset is `val/100000`. The current code divides the raw val by 100000 at read time (`colors.ts:80-85`) and adds to a 0..1 L, which is consistent for the 0..1 scale; double-check this scale equivalence when adding sat/hue mods.
  Source: https://python-pptx.readthedocs.io/en/latest/dev/analysis/txt-font-color.html
- DrawingML tint/shade (linear-RGB) is a DIFFERENT algorithm from SpreadsheetML/WordprocessingML tint/shade (HSL luminance). Keep `src/pptx/colors.ts` and `src/xlsx/colors.ts` algorithmically distinct; do not share `applyMods`.
  Source: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.color
- Gradient geometry: stop `pos` is ST_PositiveFixedPercentage in 1000ths of a percent (divide by 1000 for CSS %). `a:lin@ang` is 60000ths of a degree clockwise from 3 o'clock; CSS angle = `(ang/60000 + 90) mod 360`. `path=circle|rect` -> radial-gradient with focus from `a:fillToRect`; `path=shape` has no CSS equivalent (fall back to raster/PDF).
  Source: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill.angle
- Pattern fills: `a:pattFill@prst` is one of ~54 presets (pct5..pct90 densities, horz/vert/diag families, grids, etc.) with `a:fgClr` (hatch) over `a:bgClr` (field); default fg black, bg white, prst pct5. Map each to a small SVG `<pattern>` tile.
  Source: https://c-rex.net/samples/ooxml/e1/Part3/OOXML_P3_Primer_Pattern_topic_ID0E53OO.html
- blipFill: `a:srcRect` (l/t/r/b crop in 1000ths of a percent) crops the source; `a:stretch`+`a:fillRect` (stretch) vs `a:tile` (repeat with tx/ty/sx/sy/algn) decides fit.
  Source: https://ooxml.info/docs/20/20.5/20.5.2/20.5.2.2/
- Preset geometry: 187 `prstGeom` shapes are guide/formula programs over a 0..100000 coordinate space (avLst, gdLst, pathLst with moveTo/lnTo/arcTo/quadBezTo/cubicBezTo/close; RPN fmla ops `*/ +- +/ val abs sqrt min max sin cos tan at2 mod pin`; tokens w,h,ss,ls,hc,vc,wd2,hd2, angle constants). arcTo (center+start+sweep in 60000ths deg) converts to SVG elliptical-arc A. Implement the common ~20-40 first; route the rest to PDF.
  Source: https://learn.microsoft.com/en-us/archive/blogs/openspecification/how-to-use-the-presetshapedefinitions-xml-file-and-fun-with-drawingml

### 4.7 Font substitution (shared by both PPTX paths)

Calibri->Carlito and Cambria->Caladea are metric-compatible (identical advance widths, preserve line breaks). "Aptos Display" is a legacy family name; CoreText/LibreOffice match the typographic family "Aptos", so map "Aptos Display"->"Aptos". The `FONT_SUBSTITUTIONS` table in `src/officeToPdf.ts:103-108` is correct; mirror it in the native pptx renderer's CSS fallback chain.
Source: https://ask.libreoffice.org/t/using-the-headless-command-line-tool-soffice-to-convert-from-powerpoint-to-pdf/13374

### 4.8 LibreOffice headless PDF export filters

- Base: `soffice --headless --convert-to pdf:FILTER:JSON --outdir DIR input`. Filters: `writer_pdf_Export` (docx), `calc_pdf_Export` (xlsx), `impress_pdf_Export` (pptx).
- `SinglePageSheets:true` (calc): one un-paginated page per sheet sized to content; "ignores each sheet's paper size, print ranges and shown/hidden status." This is the fix for the cutoff and for wide-sheet pagination. Verified empirically (default clips to print area; SinglePageSheets shows the full range).
- Other useful options: `PageRange:"1"` (first page/slide for a cover), `ExportBookmarks:true` (outline for pdf.js), `SelectPdfVersion`, `Quality`.
  Source: https://help.libreoffice.org/latest/en-US/text/shared/guide/pdf_params.html
- There is no soffice CLI fit-to-width flag; fit-to-page must be set in the document, which is why `SinglePageSheets` (a filter option) is the right lever.
  Source: https://ask.libreoffice.org/t/how-to-set-fit-to-page-width-when-convert-calc-file-to-pdf-in-headless-mode/50310

### 4.9 docx-preview options (WordOptions defaults)

`className='docx'`, `inWrapper=true`, `hideWrapperOnPrint=false`, `breakPages=true`, `ignoreLastRenderedPageBreak` (library default true; this plugin correctly sets it false to honor Word's hints), `experimental=false` (plugin sets true for tab stops), `useBase64URL=false` (plugin sets true for Electron), `renderHeaders/Footers/Footnotes/Endnotes=true`, `renderChanges=false`, `renderComments=false`. Only `renderAsync` is a stable API; pin the version since fidelity tracks minor releases.
Source: https://github.com/VolodymyrBaydalka/docxjs
