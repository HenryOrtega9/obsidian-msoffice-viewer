import type ExcelJS from "exceljs";
import type { GridContext } from "./grid";

// Apply Excel-style frozen panes by promoting the relevant rows/cols to
// position:sticky. Keeps the single-table structure so merges, fills, and
// later overlay layers (images, charts) all stay in sync.
export function applyFrozenPanes(ws: ExcelJS.Worksheet, ctx: GridContext): void {
  const views = (ws as unknown as { views?: Array<{ state?: string; xSplit?: number; ySplit?: number }> }).views;
  const view = Array.isArray(views) ? views[0] : undefined;
  if (!view || view.state !== "frozen") return;

  const xSplit = view.xSplit ?? 0; // cols frozen at left
  const ySplit = view.ySplit ?? 0; // rows frozen at top
  if (xSplit <= 0 && ySplit <= 0) return;

  // The thead is already sticky-top from the base CSS. Frozen body rows must
  // start beneath it, so use the thead's rendered height as the baseline.
  const theadHeight = ctx.theadEl.offsetHeight || 0;
  const rowHdrWidth = ctx.rowHdrColPx; // typically 40

  if (ySplit > 0) {
    for (let r = 1; r <= ySplit && r <= ctx.lastRow; r++) {
      const top = theadHeight + ctx.cumulativeRowPx[r - 1];
      const rowHdr = ctx.rowHdrTh.get(r);
      if (rowHdr) {
        rowHdr.style.position = "sticky";
        rowHdr.style.top = `${top}px`;
        // Row-header column already has left:0; stay above its z-index.
        rowHdr.style.zIndex = "3";
      }
      for (let c = 1; c <= ctx.lastCol; c++) {
        const td = ctx.cellMap.get(`${r}:${c}`);
        if (!td) continue;
        td.style.position = "sticky";
        td.style.top = `${top}px`;
        td.style.zIndex = c <= xSplit ? "3" : "2";
        ensureOpaqueBackground(td);
      }
    }
  }

  if (xSplit > 0) {
    for (let c = 1; c <= xSplit && c <= ctx.lastCol; c++) {
      const left = rowHdrWidth + ctx.cumulativeColPx[c - 1];
      for (let r = 1; r <= ctx.lastRow; r++) {
        const td = ctx.cellMap.get(`${r}:${c}`);
        if (!td) continue;
        // Frozen-row cells were already set above; just add left + bump z if needed.
        td.style.position = "sticky";
        td.style.left = `${left}px`;
        if (r > ySplit) {
          td.style.zIndex = "2";
        }
        ensureOpaqueBackground(td);
      }
    }
  }
}

function ensureOpaqueBackground(td: HTMLTableCellElement): void {
  // Sticky cells without an Excel fill would show scrolled content behind them.
  // The base .docx-claude-xlsx-table td rule sets background: #ffffff, but inline
  // styles only override when present — guarantee it for sticky cells.
  if (!td.style.backgroundColor) td.style.backgroundColor = "#ffffff";
}
