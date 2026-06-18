import type ExcelJS from "exceljs";
import type { MergeRect } from "../merges";
import type { GridContext } from "../grid";
import { type CfLocks, lockSetFor, mergeStyleIntoElement } from "./applyStyle";
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
  locks: CfLocks,
  theme?: readonly string[],
): void {
  forEachCellInRanges(ranges, (r, c) => {
    const key = `${r}:${c}`;
    if (stopped.has(key)) return;
    if (!predicate(r, c)) return;
    const td = ctx.cellMap.get(key);
    if (td) mergeStyleIntoElement(td, style, theme, lockSetFor(locks, key));
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
  locks: CfLocks,
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
      case "greaterThanOrEqual":
        return n != null && a != null && n >= a;
      case "lessThanOrEqual":
        return n != null && a != null && n <= a;
      case "between":
        return n != null && a != null && b != null && n >= Math.min(a, b) && n <= Math.max(a, b);
      case "notBetween":
        return n != null && a != null && b != null && (n < Math.min(a, b) || n > Math.max(a, b));
      case "notEqual":
        // Negate equal's two-tier logic so text cells compare as strings too.
        if (n != null && a != null) return n !== a;
        return stringCellValue(ws, r, c) !== aStr;
      default:
        return false;
    }
  };
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, locks, theme);
}

export function applyTop10(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { rank?: number; percent?: boolean; bottom?: boolean; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  locks: CfLocks,
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
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, locks, theme);
}

export function applyAboveAverage(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { aboveAverage?: boolean; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  locks: CfLocks,
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
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, locks, theme);
}

export function applyContainsText(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: {
    type?: string;
    operator?: string;
    text?: string;
    formulae?: unknown[];
    style?: Partial<ExcelJS.Style>;
  },
  stopped: Set<string>,
  stopIfTrue: boolean,
  locks: CfLocks,
  theme?: readonly string[],
): void {
  // ExcelJS drops the OOXML @text attribute, so the needle survives only inside
  // the generated formula (SEARCH("x",..) for contains/notContains, LEFT/RIGHT
  // (..)="x" for begins/ends). Recover it from there, falling back to rule.text.
  const op = rule.operator ?? rule.type ?? "containsText";
  const needle = extractTextNeedle(op, rule.formulae, rule.text).toLowerCase();

  const predicate: Predicate = (r, c) => {
    const s = stringCellValue(ws, r, c).toLowerCase();
    switch (op) {
      case "containsText":
      case "containsString":
      case "contains":
        return needle !== "" && s.includes(needle);
      case "notContains":
      case "notContainsText":
        return needle === "" || !s.includes(needle);
      case "beginsWith":
        return needle !== "" && s.startsWith(needle);
      case "endsWith":
        return needle !== "" && s.endsWith(needle);
      case "containsBlanks":
        return s.trim() === "";
      case "notContainsBlanks":
        return s.trim() !== "";
      case "containsErrors":
        return cellIsError(ws, r, c);
      case "notContainsErrors":
        return !cellIsError(ws, r, c);
      default:
        return needle !== "" && s.includes(needle);
    }
  };
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, locks, theme);
}

// Recover the search literal. For contains/notContains ExcelJS emits
// SEARCH("needle",ref); for begins/ends it emits LEFT(..)="needle" /
// RIGHT(..)="needle". Excel doubles embedded quotes (""), so unescape them.
function extractTextNeedle(op: string, formulae: unknown[] | undefined, text?: string): string {
  if (typeof text === "string" && text !== "") return text;
  const f = String(formulae?.[0] ?? "");
  const unescape = (s: string) => s.replace(/""/g, '"');
  if (op === "beginsWith" || op === "endsWith") {
    const m = f.match(/="((?:[^"]|"")*)"\s*$/);
    if (m) return unescape(m[1]);
  }
  const m = f.match(/SEARCH\("((?:[^"]|"")*)"/i);
  if (m) return unescape(m[1]);
  return "";
}

// An error cell surfaces as { error } (literal) or { result: { error } }
// (formula). Inspect cell.value directly — extractCellValue flattens errors to
// the plain string "#N/A", indistinguishable from a typed-text "#N/A".
function cellIsError(ws: ExcelJS.Worksheet, r: number, c: number): boolean {
  const v = ws.getRow(r).getCell(c).value as unknown;
  if (v && typeof v === "object") {
    if ("error" in v) return true;
    const res = (v as { result?: unknown }).result;
    if (res && typeof res === "object" && "error" in res) return true;
  }
  return false;
}

export function applyTimePeriod(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { timePeriod?: string; style?: Partial<ExcelJS.Style> },
  stopped: Set<string>,
  stopIfTrue: boolean,
  locks: CfLocks,
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
  applyMatching(ws, ctx, ranges, predicate, rule.style, stopped, stopIfTrue, locks, theme);
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
