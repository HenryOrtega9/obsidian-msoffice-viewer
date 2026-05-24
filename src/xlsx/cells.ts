import type ExcelJS from "exceljs";
import { format as numfmtFormat, dateToSerial } from "numfmt";
import { ExcelColorRef, resolveExcelColor } from "./colors";
import { warn } from "./warn";
import { RichTextRun, renderRichTextRuns } from "./richText";
import {
  InternalLinkHandler,
  SheetHyperlink,
  getCellHyperlink,
  wrapInHyperlink,
} from "./hyperlinks";
import { addNoteMarker, extractCellNote } from "./notes";

export interface RenderCellOpts {
  sheetHyperlinks: Map<string, SheetHyperlink>;
  popoverHost: HTMLElement;
  onInternalLink?: InternalLinkHandler;
  theme?: readonly string[];
}

// Orchestrator: fills a <td> with content (rich text spans or plain text),
// applies inline styling, then layers hyperlinks and note markers on top.
export function renderCellInto(
  cell: ExcelJS.Cell,
  td: HTMLTableCellElement,
  opts: RenderCellOpts,
): void {
  const runs = extractRichTextRuns(cell);
  if (runs) {
    renderRichTextRuns(runs, td, opts.theme);
  } else {
    const text = cellText(cell);
    if (text) td.setText(text);
  }

  const style = cellInlineStyle(cell, opts.theme);
  if (style) td.setAttribute("style", style);

  const link = getCellHyperlink(cell, opts.sheetHyperlinks);
  if (link) wrapInHyperlink(td, link, { onInternal: opts.onInternalLink });

  const note = extractCellNote(cell);
  if (note) addNoteMarker(td, note, opts.popoverHost);
}

function extractRichTextRuns(cell: ExcelJS.Cell): RichTextRun[] | null {
  const v = cell.value as unknown;
  if (!v || typeof v !== "object") return null;
  const obj = v as { richText?: unknown; text?: unknown };
  if (Array.isArray(obj.richText)) return obj.richText as RichTextRun[];
  // Hyperlink cells may nest rich text inside .text
  if (obj.text && typeof obj.text === "object") {
    const inner = obj.text as { richText?: unknown };
    if (Array.isArray(inner.richText)) return inner.richText as RichTextRun[];
  }
  return null;
}

export function cellText(cell: ExcelJS.Cell): string {
  const raw = extractCellValue(cell);
  if (raw == null) return "";
  if (typeof raw === "string") return raw;
  const fmt = cell.numFmt;
  if (typeof raw === "boolean") {
    // Workbooks can format booleans via custom codes ("Yes";"No"). Try the
    // numfmt path first; fall back to plain TRUE/FALSE.
    if (fmt && fmt !== "General") {
      try {
        return numfmtFormat(fmt, raw ? 1 : 0);
      } catch (e) {
        warn("numfmt", e, { fmt, raw, kind: "boolean" });
      }
    }
    return raw ? "TRUE" : "FALSE";
  }
  if (raw instanceof Date) {
    if (fmt && fmt !== "General") {
      try {
        return numfmtFormat(fmt, dateToSerial(raw));
      } catch (e) {
        warn("numfmt", e, { fmt, raw, kind: "date" });
      }
    }
    return raw.toLocaleDateString();
  }
  // Number
  if (!fmt || fmt === "General" || fmt === "@") {
    return formatGeneralNumber(raw);
  }
  try {
    return numfmtFormat(fmt, raw);
  } catch (e) {
    warn("numfmt", e, { fmt, raw, kind: "number" });
    return formatGeneralNumber(raw);
  }
}

export function extractCellValue(
  cell: ExcelJS.Cell,
): number | string | boolean | Date | null {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "number" || typeof v === "string" || typeof v === "boolean")
    return v;
  if (v instanceof Date) return v;
  if (typeof v !== "object") return null;

  const obj = v as unknown as Record<string, unknown>;

  if (Array.isArray(obj.richText)) {
    return (obj.richText as { text?: string }[])
      .map((rt) => rt.text ?? "")
      .join("");
  }
  // Formula cells: { formula, result } or { sharedFormula, result }
  if ("formula" in obj || "sharedFormula" in obj) {
    return resultValue(obj.result);
  }
  // Hyperlink: { text, hyperlink }
  if ("hyperlink" in obj) {
    if (typeof obj.text === "string") return obj.text;
    // Nested rich text under hyperlink: flatten to string.
    if (obj.text && typeof obj.text === "object") {
      const inner = obj.text as { richText?: { text?: string }[] };
      if (Array.isArray(inner.richText)) {
        return inner.richText.map((rt) => rt.text ?? "").join("");
      }
    }
    return String(obj.hyperlink ?? "");
  }
  // Error cell: { error: '#N/A' }
  if ("error" in obj) {
    return String(obj.error);
  }
  if (typeof obj.text === "string") return obj.text;
  return null;
}

function resultValue(
  r: unknown,
): number | string | boolean | Date | null {
  if (r == null) return null;
  if (typeof r === "number" || typeof r === "string" || typeof r === "boolean")
    return r;
  if (r instanceof Date) return r;
  if (typeof r === "object" && r !== null && "error" in r) {
    return String((r as { error: string }).error);
  }
  return String(r);
}

export function formatGeneralNumber(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  if (Number.isInteger(n)) return String(n);
  // Trim FP artifacts (e.g. 22045.858999999997 → 22045.859). Excel's General
  // format shows up to ~11 significant digits, so ~10 decimals is plenty.
  const rounded = Math.round(n * 1e10) / 1e10;
  return String(rounded);
}

export function cellInlineStyle(cell: ExcelJS.Cell, theme?: readonly string[]): string {
  const parts: string[] = [];

  const font = cell.font as Partial<ExcelJS.Font> | undefined;
  if (font) {
    if (font.bold) parts.push("font-weight: 600");
    if (font.italic) parts.push("font-style: italic");
    if (font.underline) parts.push("text-decoration: underline");
    if (typeof font.size === "number") parts.push(`font-size: ${font.size}pt`);
    if (font.name) parts.push(`font-family: ${quoteFont(font.name)}`);
    const colorCss = resolveExcelColor(font.color as ExcelColorRef | undefined, theme);
    if (colorCss) parts.push(`color: ${colorCss}`);
  }

  const fill = cell.fill as ExcelJS.Fill | undefined;
  if (fill && fill.type === "pattern") {
    // For solid patterns fgColor is the fill. For non-solid patterns we use
    // fgColor as the primary color (good enough for a viewer) and fall back
    // to bgColor if fgColor is missing.
    const fgRef = (fill as { fgColor?: ExcelColorRef }).fgColor;
    const bgRef = (fill as { bgColor?: ExcelColorRef }).bgColor;
    const css = resolveExcelColor(fgRef, theme) ?? resolveExcelColor(bgRef, theme);
    if (css) parts.push(`background-color: ${css}`);
  }

  const border = cell.border as Partial<ExcelJS.Borders> | undefined;
  if (border) {
    if (border.top?.style) parts.push(`border-top: ${borderCss(border.top, theme)}`);
    if (border.bottom?.style) parts.push(`border-bottom: ${borderCss(border.bottom, theme)}`);
    if (border.left?.style) parts.push(`border-left: ${borderCss(border.left, theme)}`);
    if (border.right?.style) parts.push(`border-right: ${borderCss(border.right, theme)}`);
  }

  const align = cell.alignment as Partial<ExcelJS.Alignment> | undefined;
  if (align) {
    if (align.horizontal) parts.push(`text-align: ${align.horizontal}`);
    if (align.vertical) {
      // ExcelJS exposes "top" | "bottom" | "justify" | "distributed" — older
      // type defs sometimes include "middle"; treat anything center-ish as
      // vertical-align: middle.
      const raw = align.vertical as string;
      const v = raw === "middle" || raw === "center" ? "middle" : raw;
      parts.push(`vertical-align: ${v}`);
    }
    if (align.wrapText) parts.push("white-space: normal");
  }

  return parts.join("; ");
}

function quoteFont(name: string): string {
  return /^[A-Za-z0-9_-]+$/.test(name) ? name : `"${name.replace(/"/g, "")}"`;
}

function borderCss(b: Partial<ExcelJS.Border>, theme?: readonly string[]): string {
  const style = b.style ?? "thin";
  const width = style === "thick" || style === "double" ? "2px" : "1px";
  const css = style === "double" ? "double" : "solid";
  const colorCss = resolveExcelColor(b.color as ExcelColorRef | undefined, theme) ?? "#d4d4d4";
  return `${width} ${css} ${colorCss}`;
}
