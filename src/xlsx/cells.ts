import type ExcelJS from "exceljs";
import {
  format as numfmtFormat,
  formatColor as numfmtFormatColor,
  dateToSerial,
  isDateFormat,
} from "numfmt";
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
    if (text) {
      // Wrap super/subscript text in semantic <sup>/<sub> so it actually
      // renders raised/lowered (and smaller, via the UA stylesheet) without
      // colliding with the cell's vertical-align.
      const va = (cell.font as Partial<ExcelJS.Font> | undefined)?.vertAlign;
      if (va === "superscript") td.createEl("sup").setText(text);
      else if (va === "subscript") td.createEl("sub").setText(text);
      else td.setText(text);
    }
  }

  const style = cellInlineStyle(cell, opts.theme);
  if (style) td.setAttribute("style", style);

  applyCellRotation(cell, td);

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
    // ExcelJS encodes the spreadsheet wall-clock time in the Date's UTC
    // fields; ignoreTimezone keeps dateToSerial from re-reading it via the
    // host's local timezone (which would skew the serial by the UTC offset).
    const serial = dateToSerial(raw, { ignoreTimezone: true }) as number;
    if (fmt && fmt !== "General") {
      try {
        return numfmtFormat(fmt, serial);
      } catch (e) {
        warn("numfmt", e, { fmt, raw, kind: "date" });
      }
    }
    // No/General format: Excel shows a fixed short date (plus time when the
    // serial carries a fractional day), not the platform locale's rendering.
    try {
      const defFmt = serial % 1 !== 0 ? "m/d/yyyy h:mm" : "m/d/yyyy";
      return numfmtFormat(defFmt, serial);
    } catch (e) {
      warn("numfmt", e, { fmt: "General", raw, kind: "date" });
    }
    return raw.toLocaleDateString();
  }
  // Number. "@" (Text) shows a number as its General string, so both routes
  // share numfmt's General formatter.
  if (!fmt || fmt === "General" || fmt === "@") {
    return formatGeneralNumber(raw);
  }
  try {
    return numfmtFormat(fmt, adjustSerialForDate1904(cell, fmt, raw));
  } catch (e) {
    warn("numfmt", e, { fmt, raw, kind: "number" });
    return formatGeneralNumber(raw);
  }
}

// Days between the 1900 and 1904 Excel epochs. numfmt only understands the
// 1900 date system, so serials from a 1904 workbook shift forward before
// formatting. Only applies to raw numeric serials under a date format —
// ExcelJS already converts 1904 serials when it surfaces Date objects.
const DATE_1904_OFFSET = 1462;

function adjustSerialForDate1904(
  cell: ExcelJS.Cell,
  fmt: string,
  serial: number,
): number {
  if (!isDateFormat(fmt)) return serial;
  const wb = cell.worksheet?.workbook as
    | { properties?: { date1904?: boolean } }
    | undefined;
  return wb?.properties?.date1904 ? serial + DATE_1904_OFFSET : serial;
}

// A number format can carry a section color ("[Red]") or conditional-section
// color ("[<50][Red]..."). Resolve it against the cell's numeric value so it
// can override the static font color, matching how Excel paints the cell.
// Returns null for non-numeric values, the General/text format, or when the
// active section specifies no color.
export function numberFormatColor(cell: ExcelJS.Cell): string | null {
  const fmt = cell.numFmt;
  if (!fmt || fmt === "General" || fmt === "@") return null;
  const raw = extractCellValue(cell);
  let serial: number;
  if (typeof raw === "number") serial = raw;
  // ignoreTimezone: ExcelJS dates carry the wall-clock in UTC fields.
  else if (raw instanceof Date) serial = dateToSerial(raw, { ignoreTimezone: true }) as number;
  else return null;
  try {
    const color = numfmtFormatColor(fmt, serial, { indexColors: true, throws: false });
    return typeof color === "string" ? color : null;
  } catch {
    return null;
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
  // numfmt's General formatter carries Excel's 11-significant-digit and
  // scientific-notation switchover rules.
  try {
    return numfmtFormat("General", n);
  } catch (e) {
    warn("numfmt", e, { fmt: "General", raw: n, kind: "number" });
  }
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
    // Excel "bold" is weight 700, not 600.
    if (font.bold) parts.push("font-weight: bold");
    if (font.italic) parts.push("font-style: italic");
    // strike and underline both live on text-decoration; merge them so one
    // doesn't clobber the other. "double" underline keeps the double style.
    const decorations: string[] = [];
    if (font.underline) {
      decorations.push(font.underline === "double" ? "underline double" : "underline");
    }
    if (font.strike) decorations.push("line-through");
    if (decorations.length) parts.push(`text-decoration: ${decorations.join(" ")}`);
    if (typeof font.size === "number") parts.push(`font-size: ${font.size}pt`);
    if (font.name) parts.push(`font-family: ${quoteFont(font.name)}`);
    // super/subscript is applied by wrapping the text in <sup>/<sub> in
    // renderCellInto, NOT a cell-level vertical-align (which would collide with
    // and be clobbered by the alignment.vertical declaration below).
    const colorCss = resolveExcelColor(font.color as ExcelColorRef | undefined, theme);
    if (colorCss) parts.push(`color: ${colorCss}`);
  }

  // Number-format-driven color (e.g. "#,##0;[Red]-#,##0" shows negatives red).
  // Applied after the static font color so it overrides it, matching Excel.
  const fmtColor = numberFormatColor(cell);
  if (fmtColor) parts.push(`color: ${fmtColor}`);

  const fill = cell.fill as ExcelJS.Fill | undefined;
  if (fill) {
    const fillCss = fillToCss(fill, theme);
    if (fillCss) parts.push(fillCss);
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
    // Indent levels (~3 char-widths each in Excel). Pivot tables encode their
    // grouped row-label hierarchy purely as indent, so honoring it restores
    // the visual nesting. Emit a full-side padding (overriding only that side's
    // base cell pad) for positive indent; leave indent-0 cells untouched.
    const indent = typeof align.indent === "number" ? align.indent : 0;
    if (indent > 0) {
      const rtl = (align as { readingOrder?: unknown }).readingOrder === "rtl";
      const side = align.horizontal === "right" || rtl ? "padding-right" : "padding-left";
      parts.push(`${side}: ${6 + indent * 9}px`);
    }
    if (align.wrapText) parts.push("white-space: normal");
  }

  return parts.join("; ");
}

function quoteFont(name: string): string {
  return /^[A-Za-z0-9_-]+$/.test(name) ? name : `"${name.replace(/"/g, "")}"`;
}

// Approximate fraction of a cell covered by the foreground color for each
// non-solid OOXML fill pattern. Used to blend fg over bg so a sparse pattern
// (e.g. gray125) reads as a light tint instead of a saturated solid.
const PATTERN_DENSITY: Record<string, number> = {
  gray0625: 0.0625,
  gray125: 0.125,
  lightGray: 0.25,
  lightHorizontal: 0.25,
  lightVertical: 0.25,
  lightDown: 0.25,
  lightUp: 0.25,
  lightGrid: 0.4,
  lightTrellis: 0.4,
  mediumGray: 0.5,
  darkHorizontal: 0.5,
  darkVertical: 0.5,
  darkDown: 0.5,
  darkUp: 0.5,
  darkGray: 0.75,
  darkGrid: 0.75,
  darkTrellis: 0.85,
};

// Parse "#rgb"/"#rrggbb" to [r,g,b]; returns null for non-hex inputs (e.g.
// rgb()/named colors) so callers can fall back gracefully.
function parseHex(css: string): [number, number, number] | null {
  const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(css.trim());
  if (!m) return null;
  let h = m[1];
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

// Linear blend of `over` onto `base` by fraction t (0..1).
function blendHex(base: string, over: string, t: number): string | null {
  const a = parseHex(base);
  const b = parseHex(over);
  if (!a || !b) return null;
  const mix = (i: number) => Math.round(a[i] * (1 - t) + b[i] * t);
  const hx = (n: number) => n.toString(16).padStart(2, "0");
  return `#${hx(mix(0))}${hx(mix(1))}${hx(mix(2))}`;
}

// Turn an ExcelJS fill into a CSS background declaration. Solid patterns paint
// fgColor; non-solid patterns blend fg over bg by approximate density; gradient
// fills become a linear/radial CSS gradient from the stop list.
function fillToCss(fill: ExcelJS.Fill, theme?: readonly string[]): string | null {
  if (fill.type === "pattern") {
    const pat = (fill as { pattern?: string }).pattern;
    if (!pat || pat === "none") return null;
    const fgRef = (fill as { fgColor?: ExcelColorRef }).fgColor;
    const bgRef = (fill as { bgColor?: ExcelColorRef }).bgColor;
    const fg = resolveExcelColor(fgRef, theme);
    const bg = resolveExcelColor(bgRef, theme);
    if (pat === "solid") {
      const css = fg ?? bg;
      return css ? `background-color: ${css}` : null;
    }
    const density = PATTERN_DENSITY[pat] ?? 0.5;
    const blended = blendHex(bg ?? "#ffffff", fg ?? "#000000", density);
    const css = blended ?? fg ?? bg;
    return css ? `background-color: ${css}` : null;
  }
  if (fill.type === "gradient") {
    const g = fill as {
      gradient?: string;
      degree?: number;
      stops?: { position?: number; color?: ExcelColorRef }[];
    };
    const stops = (g.stops ?? [])
      .map((s) => {
        const c = resolveExcelColor(s.color, theme);
        if (!c) return null;
        return `${c} ${Math.round((s.position ?? 0) * 100)}%`;
      })
      .filter((s): s is string => s != null);
    if (stops.length < 2) {
      return stops.length === 1
        ? `background-color: ${stops[0].split(" ")[0]}`
        : null;
    }
    if (g.gradient === "path") {
      return `background-image: radial-gradient(${stops.join(", ")})`;
    }
    const deg = typeof g.degree === "number" ? g.degree : 90;
    return `background-image: linear-gradient(${deg}deg, ${stops.join(", ")})`;
  }
  return null;
}

// Excel text rotation -> CSS. Positive Excel rotation is counterclockwise (CSS
// rotate() is clockwise, so the sign flips); 255/"vertical" is stacked upright
// text. Applied to a wrapper span so the cell box keeps laying out normally.
function cellRotationStyle(cell: ExcelJS.Cell): string | null {
  const align = cell.alignment as Partial<ExcelJS.Alignment> | undefined;
  if (!align) return null;
  const rot = align.textRotation as number | "vertical" | undefined;
  if (rot == null) return null;
  if (rot === "vertical" || rot === 255) {
    return "writing-mode: vertical-rl; text-orientation: upright; white-space: nowrap";
  }
  if (typeof rot === "number" && rot !== 0) {
    return `display: inline-block; white-space: nowrap; transform: rotate(${-rot}deg)`;
  }
  return null;
}

// Wrap a rotated cell's content in a span carrying the transform, leaving the
// <td> itself unrotated so column widths and note markers stay put.
function applyCellRotation(cell: ExcelJS.Cell, td: HTMLTableCellElement): void {
  const rotStyle = cellRotationStyle(cell);
  if (!rotStyle) return;
  const kids = Array.from(td.childNodes);
  if (!kids.length) return;
  const inner = td.createSpan();
  inner.setAttribute("style", rotStyle);
  for (const k of kids) inner.appendChild(k);
}

// OOXML border style -> CSS {width, style}. Excel's automatic (unspecified)
// border color is black, not a light gray.
const BORDER_STYLE_CSS: Record<string, { width: string; css: string }> = {
  hair: { width: "1px", css: "solid" },
  thin: { width: "1px", css: "solid" },
  dotted: { width: "1px", css: "dotted" },
  dashed: { width: "1px", css: "dashed" },
  dashDot: { width: "1px", css: "dashed" },
  dashDotDot: { width: "1px", css: "dashed" },
  slantDashDot: { width: "1px", css: "dashed" },
  medium: { width: "2px", css: "solid" },
  mediumDashed: { width: "2px", css: "dashed" },
  mediumDashDot: { width: "2px", css: "dashed" },
  mediumDashDotDot: { width: "2px", css: "dashed" },
  thick: { width: "3px", css: "solid" },
  double: { width: "3px", css: "double" },
};

function borderCss(b: Partial<ExcelJS.Border>, theme?: readonly string[]): string {
  const style = b.style ?? "thin";
  const { width, css } = BORDER_STYLE_CSS[style] ?? BORDER_STYLE_CSS.thin;
  const colorCss = resolveExcelColor(b.color as ExcelColorRef | undefined, theme) ?? "#000000";
  return `${width} ${css} ${colorCss}`;
}
