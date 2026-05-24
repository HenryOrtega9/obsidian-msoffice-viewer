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
// the maximum digit width." 7px per char is a reasonable approximation for
// Calibri 11 on a typical Mac.
export const COL_WIDTH_PX_PER_CH = 7;
export const DEFAULT_COL_WIDTH_CH = 8.43;
export const POINTS_TO_PX = 4 / 3;
export const DEFAULT_ROW_HEIGHT_PX = 20;
export const ROW_HDR_COL_WIDTH_PX = 40;

// Excel paints an empty grid canvas beyond the data. Pad the rendered range
// with trailing empty rows/columns so a small sheet doesn't look chopped off
// at its last value. Math.max keeps large sheets from being inflated.
const MIN_GRID_COLS = 16;
const MIN_GRID_ROWS = 50;
const COL_PAD = 2;
const ROW_PAD = 8;

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
  let lastCol = Math.max(1, ws.actualColumnCount || ws.columnCount || 1);
  let lastRow = Math.max(1, ws.actualRowCount || ws.rowCount || 1);
  for (const m of merges) {
    if (m.right > lastCol) lastCol = m.right;
    if (m.bottom > lastRow) lastRow = m.bottom;
  }

  // Pad with empty trailing rows/columns for a spreadsheet-like canvas.
  lastCol = Math.max(lastCol + COL_PAD, MIN_GRID_COLS);
  lastRow = Math.max(lastRow + ROW_PAD, MIN_GRID_ROWS);

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

  const cumulativeColPx: number[] = [0]; // cumulativeColPx[0] = 0 (offset before col 1)
  for (let c = 1; c <= lastCol; c++) {
    const col = ws.getColumn(c);
    const widthCh = col.width ?? DEFAULT_COL_WIDTH_CH;
    const widthPx = Math.round(widthCh * COL_WIDTH_PX_PER_CH);
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

  const tbody = table.createEl("tbody");
  const cumulativeRowPx: number[] = [0];
  for (let r = 1; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const tr = tbody.createEl("tr");
    const heightPx = row.height
      ? Math.round(row.height * POINTS_TO_PX)
      : DEFAULT_ROW_HEIGHT_PX;
    if (row.height) {
      tr.setAttribute("style", `height: ${heightPx}px`);
    }
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
