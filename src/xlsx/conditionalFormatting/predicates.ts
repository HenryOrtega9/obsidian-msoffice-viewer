import type ExcelJS from "exceljs";
import type { MergeRect } from "../merges";
import type { GridContext } from "../grid";
import { mergeStyleIntoElement } from "./applyStyle";
import {
  collectNumericValues,
  forEachCellInRanges,
  numericCellValue,
  stringCellValue,
} from "./values";

type Predicate = (row: number, col: number) => boolean;

function applyMatching(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  predicate: Predicate,
  style: Partial<ExcelJS.Style> | undefined,
  stopped: Set<string>,
  stopIfTrue: boolean,
  theme?: readonly string[],
): void {
  forEachCellInRanges(ranges, (r, c) => {
    const key = `${r}:${c}`;
    if (stopped.has(key)) return;
    if (!predicate(r, c)) return;
    const td = ctx.cellMap.get(key);
    if (td) mergeStyleIntoElement(td, style, theme);
    if (stopIfTrue) stopped.add(key);
  });
}

export function applyCellIs(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { operator?: string; formulae?: unknown[]; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  theme?: readonly string[],
): void {
  const f = rule.formulae ?? [];
  const a = toNum(f[0]);
  const b = toNum(f[1]);
  const aStr = f[0] != null ? String(f[0]).replace(/^"|"$/g, "") : "";

  const predicate: Predicate = (r, c) => {
    const n = numericCellValue(ws, r, c);
    switch (rule.operator) {
      case "equal":
        if (n != null && a != null) return n === a;
        return stringCellValue(ws, r, c) === aStr;
      case "greaterThan":
        return n != null && a != null && n > a;
      case "lessThan":
        return n != null && a != null && n < a;
      case "between":
        return n != null && a != null && b != null && n >= Math.min(a, b) && n <= Math.max(a, b);
      default:
        return false;
    }
  };
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, theme);
}

export function applyTop10(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { rank?: number; percent?: boolean; bottom?: boolean; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  theme?: readonly string[],
): void {
  const values = collectNumericValues(ws, ranges);
  if (values.length === 0) return;
  const rank = rule.rank ?? 10;
  let n = rule.percent ? Math.ceil((rank / 100) * values.length) : rank;
  n = Math.max(1, Math.min(values.length, n));
  const sorted = [...values].sort((x, y) => (rule.bottom ? x - y : y - x));
  const threshold = sorted[n - 1];

  const predicate: Predicate = (r, c) => {
    const v = numericCellValue(ws, r, c);
    if (v == null) return false;
    return rule.bottom ? v <= threshold : v >= threshold;
  };
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, theme);
}

export function applyAboveAverage(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { aboveAverage?: boolean; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  theme?: readonly string[],
): void {
  const values = collectNumericValues(ws, ranges);
  if (values.length === 0) return;
  const mean = values.reduce((s, v) => s + v, 0) / values.length;
  const above = rule.aboveAverage !== false;

  const predicate: Predicate = (r, c) => {
    const v = numericCellValue(ws, r, c);
    if (v == null) return false;
    return above ? v > mean : v < mean;
  };
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, theme);
}

export function applyContainsText(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { operator?: string; text?: string; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  theme?: readonly string[],
): void {
  const needle = (rule.text ?? "").toLowerCase();
  const predicate: Predicate = (r, c) => {
    const s = stringCellValue(ws, r, c).toLowerCase();
    switch (rule.operator) {
      case "containsText":
        return needle !== "" && s.includes(needle);
      case "containsBlanks":
        return s.trim() === "";
      case "notContainsBlanks":
        return s.trim() !== "";
      default:
        return needle !== "" && s.includes(needle);
    }
  };
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, theme);
}

export function applyTimePeriod(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { timePeriod?: string; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  theme?: readonly string[],
): void {
  const now = new Date();
  const startOfDay = (d: Date) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const today = startOfDay(now);
  const day = 86400000;

  const predicate: Predicate = (r, c) => {
    const cell = ws.getRow(r).getCell(c);
    const v = cell.value;
    if (!(v instanceof Date)) return false;
    const t = startOfDay(v);
    switch (rule.timePeriod) {
      case "today": return t === today;
      case "yesterday": return t === today - day;
      case "tomorrow": return t === today + day;
      case "last7Days": return t <= today && t >= today - 6 * day;
      case "lastWeek": return t < today - dayOfWeek(now) * day && t >= today - (dayOfWeek(now) + 7) * day;
      case "thisWeek": return t >= today - dayOfWeek(now) * day && t <= today + (6 - dayOfWeek(now)) * day;
      case "nextWeek": return t > today + (6 - dayOfWeek(now)) * day && t <= today + (13 - dayOfWeek(now)) * day;
      case "lastMonth": return sameMonth(v, addMonths(now, -1));
      case "thisMonth": return sameMonth(v, now);
      case "nextMonth": return sameMonth(v, addMonths(now, 1));
      default: return false;
    }
  };
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, theme);
}

function dayOfWeek(d: Date): number {
  return d.getDay();
}

function addMonths(d: Date, n: number): Date {
  return new Date(d.getFullYear(), d.getMonth() + n, 1);
}

function sameMonth(a: Date, b: Date): boolean {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
}

function toNum(x: unknown): number | null {
  if (typeof x === "number") return x;
  if (typeof x === "string") {
    const n = parseFloat(x);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}
