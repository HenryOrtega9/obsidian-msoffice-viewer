import { Notice } from "obsidian";
import type ExcelJS from "exceljs";
import { renderCellInto } from "./cells";
import {
  MergeRect,
  collectMerges,
  computeMergeSkipMap,
  colLetter,
} from "./merges";
import {
  InternalLinkHandler,
  collectSheetHyperlinks,
} from "./hyperlinks";

// Excel column width units are "characters of the default font, rendered with
// the maximum digit width" (MDW). For Calibri 11 at 96 DPI, MDW = 7px. exceljs
// col.width is the stored XLSX width attribute in MDW units (padding-inclusive),
// not the 8.43 UI character value. DEFAULT_COL_WIDTH_CH is kept only as a
// documented reference to that UI value.
export const MDW = 7;
export const DEFAULT_COL_WIDTH_CH = 8.43;
export const POINTS_TO_PX = 4 / 3;
// Default column width in pixels when no width is stored (stored 9.140625 maps
// to 64px via storedWidthToPx; see the assertion below).
const DEFAULT_COL_WIDTH_PX = 64;
// 15pt is Excel's default row height for Calibri 11; 15 * 4/3 = 20px at 96 DPI.
export const DEFAULT_ROW_HEIGHT_PX = Math.round(15 * POINTS_TO_PX);
export const ROW_HDR_COL_WIDTH_PX = 40;

// ECMA-376 stored-width to pixels. The Trunc(128 / MDW) term self-includes the
// per-column padding, so do NOT add a separate +5. Worked example:
// 9.140625 -> Trunc((256*9.140625 + Trunc(128/7)) / 256 * 7) = Trunc(64.4766) = 64.
function storedWidthToPx(width: number): number {
  return Math.trunc(((256 * width + Math.trunc(128 / MDW)) / 256) * MDW);
}

// Unit assertion documenting the canonical default-column mapping. If the
// formula ever drifts, this throws at module load instead of silently
// misrendering every grid.
if (storedWidthToPx(9.140625) !== DEFAULT_COL_WIDTH_PX) {
  throw new Error(
    `storedWidthToPx(9.140625) = ${storedWidthToPx(9.140625)}, expected ${DEFAULT_COL_WIDTH_PX}`,
  );
}

// Excel paints an empty grid canvas beyond the data. Pad the rendered range
// with trailing empty rows/columns so a small sheet doesn't look chopped off
// at its last value. Math.max keeps large sheets from being inflated.
const MIN_GRID_COLS = 16;
const MIN_GRID_ROWS = 50;
const COL_PAD = 2;
const ROW_PAD = 8;

// Hard ceiling on the rendered range. The grid builds one <td> per cell
// synchronously on the UI thread, so a sheet whose stored dimension or a single
// far-flung cell claims a range up to 1,048,576 x 16,384 would create millions
// of nodes and freeze (or OOM) Obsidian. These caps keep the DOM bounded; a
// genuinely larger sheet is truncated with a notice (and renders fine via the
// default PDF path regardless).
const HARD_MAX_ROWS = 5000;
const HARD_MAX_COLS = 256;

export interface GridContext {
  sheetWrapEl: HTMLElement;
  tableEl: HTMLTableElement;
  theadEl: HTMLTableSectionElement;
  tbodyEl: HTMLTableSectionElement;
  cellMap: Map<string, HTMLTableCellElement>;
  rowHdrTh: Map<number, HTMLTableCellElement>;
  // cumulativeColPx[c] = left offset (px) of column c (1-indexed)
  // cumulativeRowPx[r] = top offset (px) of row r (1-indexed, relative to tbody)
  cumulativeColPx: number[];
  cumulativeRowPx: number[];
  rowHdrColPx: number;
  lastRow: number;
  lastCol: number;
}

export interface RenderGridOpts {
  onInternalLink?: InternalLinkHandler;
  theme?: readonly string[];
}

export function renderSheetIntoGrid(
  ws: ExcelJS.Worksheet,
  container: HTMLElement,
  opts: RenderGridOpts = {},
): GridContext {
  const merges = collectMerges(ws);
  // Extend dimensions to cover every merge bound so anchor cells that sit
  // beyond actualColumnCount/actualRowCount still get rendered (and the
  // spans never overrun the colgroup).
  // Seed from the POPULATED extent (count of non-empty rows/cols).
  let lastCol = Math.max(1, ws.actualColumnCount || ws.columnCount || 1);
  let lastRow = Math.max(1, ws.actualRowCount || ws.rowCount || 1);
  // ws.dimensions reflects the stored used range; actualRowCount/actualColumnCount
  // count non-empty rows/cols and can under-report trailing format-only or
  // merge-only cells, so fold the stored bounds in — but ONLY when they sit
  // close to the populated extent. A bogus stored <dimension> or a single
  // far-flung cell (e.g. A100000 from a whole-column format) otherwise inflates
  // the range to a freezing/OOM size; in that case keep the populated extent
  // and let the hard cap below be the backstop.
  const dims = ws.dimensions as { bottom?: number; right?: number } | undefined;
  if (dims) {
    if (typeof dims.right === "number" && dims.right > lastCol && dims.right <= lastCol + COL_PAD * 4) {
      lastCol = dims.right;
    }
    if (typeof dims.bottom === "number" && dims.bottom > lastRow && dims.bottom <= lastRow + ROW_PAD * 4) {
      lastRow = dims.bottom;
    }
  }
  for (const m of merges) {
    if (m.right > lastCol) lastCol = m.right;
    if (m.bottom > lastRow) lastRow = m.bottom;
  }

  // Pad with empty trailing rows/columns for a spreadsheet-like canvas.
  lastCol = Math.max(lastCol + COL_PAD, MIN_GRID_COLS);
  lastRow = Math.max(lastRow + ROW_PAD, MIN_GRID_ROWS);

  // Backstop: never build an unbounded DOM. Truncate to the hard caps and warn
  // so data isn't silently dropped.
  if (lastRow > HARD_MAX_ROWS || lastCol > HARD_MAX_COLS) {
    new Notice(
      `This sheet is very large; showing the first ${Math.min(lastRow, HARD_MAX_ROWS)} rows × ` +
        `${Math.min(lastCol, HARD_MAX_COLS)} columns. Use "Open in Excel" for the full sheet.`,
      8000,
    );
    lastCol = Math.min(lastCol, HARD_MAX_COLS);
    lastRow = Math.min(lastRow, HARD_MAX_ROWS);
  }

  const skipMap = computeMergeSkipMap(merges);
  const mergeAnchor = new Map<string, MergeRect>();
  for (const m of merges) mergeAnchor.set(`${m.top}:${m.left}`, m);

  const sheetHyperlinks = collectSheetHyperlinks(ws);
  const cellMap = new Map<string, HTMLTableCellElement>();
  const rowHdrTh = new Map<number, HTMLTableCellElement>();

  // Positioned wrapper so the image/chart overlay layers and note popovers can
  // be absolutely positioned within the table's full (scrollable) bounds.
  const sheetWrap = container.createDiv({ cls: "docx-claude-xlsx-sheet-wrap" });
  const table = sheetWrap.createEl("table", { cls: "docx-claude-xlsx-table" });

  const colgroup = table.createEl("colgroup");
  colgroup.createEl("col", { cls: "docx-claude-xlsx-rowhdr-col" });

  // Sheet default column width (stored attribute) when present, else the
  // canonical 64px default.
  const sheetDefaultColWidth = (
    ws.properties as { defaultColWidth?: number } | undefined
  )?.defaultColWidth;
  const defaultColPx =
    typeof sheetDefaultColWidth === "number"
      ? storedWidthToPx(sheetDefaultColWidth)
      : DEFAULT_COL_WIDTH_PX;

  const cumulativeColPx: number[] = [0]; // cumulativeColPx[0] = 0 (offset before col 1)
  for (let c = 1; c <= lastCol; c++) {
    const col = ws.getColumn(c);
    // Hidden columns collapse to 0px (mirrors the hidden-row branch below):
    // table-layout:fixed + overflow:hidden clips the cell, and a 0 here keeps
    // cumulativeColPx — and thus the image/chart overlay offsets — aligned.
    const widthPx = col.hidden
      ? 0
      : typeof col.width === "number"
        ? storedWidthToPx(col.width)
        : defaultColPx;
    const colEl = colgroup.createEl("col");
    colEl.setAttribute("style", `width: ${widthPx}px`);
    cumulativeColPx.push(cumulativeColPx[c - 1] + widthPx);
  }

  const thead = table.createEl("thead");
  const hdrRow = thead.createEl("tr");
  hdrRow.createEl("th", { cls: "docx-claude-xlsx-corner" });
  for (let c = 1; c <= lastCol; c++) {
    hdrRow.createEl("th", {
      cls: "docx-claude-xlsx-colhdr",
      text: colLetter(c),
    });
  }

  // Sheet default row height (points) when present, else the 15pt-derived
  // default. defaultRowHeight is stored in points like row.height.
  const sheetDefaultRowHeight = (
    ws.properties as { defaultRowHeight?: number } | undefined
  )?.defaultRowHeight;
  const defaultRowPx =
    typeof sheetDefaultRowHeight === "number"
      ? Math.round(sheetDefaultRowHeight * POINTS_TO_PX)
      : DEFAULT_ROW_HEIGHT_PX;

  const tbody = table.createEl("tbody");
  const cumulativeRowPx: number[] = [0];
  for (let r = 1; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const tr = tbody.createEl("tr");
    // Hidden or explicit 0-height rows collapse to 0px; an unset height uses the
    // sheet default; otherwise convert the stored points. Always set the tr
    // inline height so DOM rows match cumulativeRowPx for the overlay layer.
    let heightPx: number;
    if (row.hidden || row.height === 0) {
      heightPx = 0;
    } else if (row.height == null) {
      heightPx = defaultRowPx;
    } else {
      heightPx = Math.round(row.height * POINTS_TO_PX);
    }
    tr.setAttribute("style", `height: ${heightPx}px`);
    cumulativeRowPx.push(cumulativeRowPx[r - 1] + heightPx);

    const rowHdr = tr.createEl("th", {
      cls: "docx-claude-xlsx-rowhdr",
      text: String(r),
    });
    rowHdrTh.set(r, rowHdr);

    for (let c = 1; c <= lastCol; c++) {
      const key = `${r}:${c}`;
      if (skipMap.has(key)) continue;
      const cell = row.getCell(c);
      const td = tr.createEl("td", { cls: "docx-claude-xlsx-cell" });
      const merge = mergeAnchor.get(key);
      if (merge) {
        td.colSpan = merge.right - merge.left + 1;
        td.rowSpan = merge.bottom - merge.top + 1;
      }
      renderCellInto(cell, td, {
        sheetHyperlinks,
        popoverHost: sheetWrap,
        onInternalLink: opts.onInternalLink,
        theme: opts.theme,
      });
      cellMap.set(key, td);
    }
  }

  return {
    sheetWrapEl: sheetWrap,
    tableEl: table,
    theadEl: thead,
    tbodyEl: tbody,
    cellMap,
    rowHdrTh,
    cumulativeColPx,
    cumulativeRowPx,
    rowHdrColPx: ROW_HDR_COL_WIDTH_PX,
    lastRow,
    lastCol,
  };
}
