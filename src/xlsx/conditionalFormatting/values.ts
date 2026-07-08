import type ExcelJS from "exceljs";
import type { MergeRect } from "../merges";
import { colNum } from "../merges";
import { extractCellValue, excelDateToSerial } from "../cells";

export interface Cfvo {
  type: string;
  value?: number;
}

// Parse a CF ref like "A1", "A1:B20", or "A1:A20 C1:C20" into rectangles.
export function parseCfRanges(ref: string): MergeRect[] {
  const out: MergeRect[] = [];
  for (const part of ref.split(/\s+/).filter(Boolean)) {
    const rect = parseOneRange(part);
    if (rect) out.push(rect);
  }
  return out;
}

function parseOneRange(addr: string): MergeRect | null {
  const range = addr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (range) {
    return {
      top: parseInt(range[2], 10),
      left: colNum(range[1]),
      bottom: parseInt(range[4], 10),
      right: colNum(range[3]),
    };
  }
  const single = addr.match(/^([A-Z]+)(\d+)$/);
  if (single) {
    const r = parseInt(single[2], 10);
    const c = colNum(single[1]);
    return { top: r, left: c, bottom: r, right: c };
  }
  return null;
}

export function forEachCellInRanges(
  ranges: MergeRect[],
  fn: (row: number, col: number) => void,
): void {
  for (const rect of ranges) {
    for (let r = rect.top; r <= rect.bottom; r++) {
      for (let c = rect.left; c <= rect.right; c++) {
        fn(r, c);
      }
    }
  }
}

export function numericCellValue(ws: ExcelJS.Worksheet, row: number, col: number): number | null {
  const cell = ws.getRow(row).getCell(col);
  const v = extractCellValue(cell);
  if (typeof v === "number" && Number.isFinite(v)) return v;
  // Excel compares dates as serial numbers (rule formulas like "45300" are
  // serials), not epoch milliseconds.
  if (v instanceof Date) return excelDateToSerial(v);
  return null;
}

export function stringCellValue(ws: ExcelJS.Worksheet, row: number, col: number): string {
  const cell = ws.getRow(row).getCell(col);
  const v = extractCellValue(cell);
  return v == null ? "" : String(v);
}

export function collectNumericValues(ws: ExcelJS.Worksheet, ranges: MergeRect[]): number[] {
  const values: number[] = [];
  forEachCellInRanges(ranges, (r, c) => {
    const n = numericCellValue(ws, r, c);
    if (n != null) values.push(n);
  });
  return values;
}

// Resolve a single cfvo entry to its numeric threshold against the data set.
export function cfvoToNumber(
  cfvo: Cfvo,
  sorted: number[],
  min: number,
  max: number,
): number {
  switch (cfvo.type) {
    case "min":
      return min;
    case "max":
      return max;
    // Excel's automatic data-bar bounds clamp through zero: all-positive data
    // bars grow from 0, not from the smallest value.
    case "autoMin":
      return Math.min(0, min);
    case "autoMax":
      return Math.max(0, max);
    case "num":
      return cfvo.value ?? min;
    case "percent":
      return min + ((cfvo.value ?? 0) / 100) * (max - min);
    case "percentile":
      return percentile(sorted, cfvo.value ?? 0);
    case "formula":
      // We don't evaluate formulas; treat a literal numeric value as-is,
      // otherwise fall back to min.
      return typeof cfvo.value === "number" ? cfvo.value : min;
    default:
      return min;
  }
}

export function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0;
  if (sorted.length === 1) return sorted[0];
  const rank = (Math.max(0, Math.min(100, p)) / 100) * (sorted.length - 1);
  const lo = Math.floor(rank);
  const hi = Math.ceil(rank);
  if (lo === hi) return sorted[lo];
  return sorted[lo] + (rank - lo) * (sorted[hi] - sorted[lo]);
}
