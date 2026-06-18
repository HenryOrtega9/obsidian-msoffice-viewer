import type { GridContext } from "./grid";

export const EMU_PER_PX = 9525; // 914400 EMU/inch ÷ 96 px/inch

export interface AnchorPoint {
  col: number; // 0-based column index
  colOff: number; // EMU offset within column
  row: number; // 0-based row index
  rowOff: number; // EMU offset within row
}

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

// Convert a from/to cell anchor (or from + ext size) to a pixel box within the
// sheet wrapper. Adds the row-header column width and thead height so the box
// lines up with the rendered grid coordinate space.
export function anchorRangeToBox(
  from: AnchorPoint,
  to: AnchorPoint | null,
  ext: { width: number; height: number } | null,
  ctx: GridContext,
): Box {
  const left = ctx.rowHdrColPx + cumulativeAt(ctx.cumulativeColPx, from.col) + from.colOff / EMU_PER_PX;
  const top = ctx.theadEl.offsetHeight + cumulativeAt(ctx.cumulativeRowPx, from.row) + from.rowOff / EMU_PER_PX;

  if (to) {
    const right = ctx.rowHdrColPx + cumulativeAt(ctx.cumulativeColPx, to.col) + to.colOff / EMU_PER_PX;
    const bottom = ctx.theadEl.offsetHeight + cumulativeAt(ctx.cumulativeRowPx, to.row) + to.rowOff / EMU_PER_PX;
    return { left, top, width: Math.max(1, right - left), height: Math.max(1, bottom - top) };
  }
  if (ext) {
    // ext.width/height are already pixels (ExcelJS yields cx/cy already divided
    // by EMU_PER_PX; chart anchors pass cx/cy converted to px). The from.colOff/
    // rowOff offsets above are genuine EMU and stay divided. Dividing ext again
    // collapsed one-cell-anchored images/charts to a 1px dot.
    return {
      left,
      top,
      width: Math.max(1, ext.width),
      height: Math.max(1, ext.height),
    };
  }
  return { left, top, width: 1, height: 1 };
}

// cumulative[i] is defined for 0..length-1. For indices beyond the rendered
// grid, extrapolate from the last column/row gap so out-of-range anchors land
// somewhere sensible rather than NaN.
export function cumulativeAt(cumulative: number[], idx: number): number {
  if (idx < 0) return 0;
  if (idx < cumulative.length) return cumulative[idx];
  const last = cumulative[cumulative.length - 1] ?? 0;
  const prev = cumulative[cumulative.length - 2] ?? last;
  const step = Math.max(1, last - prev);
  return last + (idx - (cumulative.length - 1)) * step;
}
