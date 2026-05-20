import { Notice, TFile } from "obsidian";
import ExcelJS from "exceljs";
import { format as numfmtFormat, dateToSerial } from "numfmt";
import { OfficeFileView } from "./OfficeFileView";
import { findSoffice } from "./officeToPdf";

export const XLSX_CLAUDE_VIEW_TYPE = "xlsx-claude-view";

interface SheetEntry {
  name: string;
  worksheet: ExcelJS.Worksheet;
}

export class XlsxPreviewView extends OfficeFileView {
  private workbook: ExcelJS.Workbook | null = null;
  private sheets: SheetEntry[] = [];
  private gridEl: HTMLElement | null = null;
  private tabsEl: HTMLElement | null = null;

  getViewType(): string {
    return XLSX_CLAUDE_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Xlsx preview";
  }

  getIcon(): string {
    return "table";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "xlsx" || extension === "xls";
  }

  protected getExternalAppLabel(): string {
    return "Open in Excel";
  }

  async onUnloadFile(file: TFile): Promise<void> {
    this.resetState();
    await super.onUnloadFile(file);
  }

  private resetState(): void {
    this.workbook = null;
    this.sheets = [];
    this.gridEl = null;
    if (this.tabsEl) {
      this.tabsEl.remove();
      this.tabsEl = null;
    }
  }

  protected async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.resetState();
    this.renderEl.empty();

    if (file.extension === "xlsx") {
      try {
        await this.renderViaExcelJsGrid(file);
        return;
      } catch (e) {
        console.error("ExcelJS grid render failed; falling back to LibreOffice PDF:", e);
        new Notice("Grid rendering failed; using PDF fallback.");
        if (this.file !== file || !this.renderEl) return;
        // resetState before empty() so a half-built tabsEl (which lives
        // outside renderEl) doesn't ghost beneath the PDF fallback.
        this.resetState();
        this.renderEl.empty();
      }
    }

    // .xls (legacy binary) or ExcelJS fallback path
    const sofficeBin = await findSoffice();
    if (sofficeBin) {
      try {
        await this.renderViaLibreOfficePdf(file, sofficeBin, file.extension);
        return;
      } catch (e) {
        console.error("LibreOffice render failed:", e);
      }
    }

    if (this.renderEl) {
      this.renderEl
        .createDiv({ cls: "docx-claude-pdf-error" })
        .setText(
          file.extension === "xls"
            ? "Couldn't render this .xls. Install LibreOffice or convert to .xlsx."
            : "Couldn't render this workbook. See console for details.",
        );
    }
  }

  private async renderViaExcelJsGrid(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    const buf = await this.app.vault.readBinary(file);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    if (this.file !== file || !this.renderEl) return;
    this.workbook = wb;

    this.sheets = [];
    wb.worksheets.forEach((ws) => {
      if (ws.state === "veryHidden") return;
      this.sheets.push({ name: ws.name, worksheet: ws });
    });

    if (this.sheets.length === 0) {
      this.renderEl
        .createDiv({ cls: "docx-claude-pdf-error" })
        .setText("This workbook has no visible sheets.");
      return;
    }

    this.gridEl = this.renderEl.createDiv({ cls: "docx-claude-xlsx-grid" });

    if (this.sheets.length > 1) {
      this.tabsEl = createDiv({ cls: "docx-claude-xlsx-tabs" });
      this.buildTabs();
      this.appendBelowRender(this.tabsEl);
    }

    this.showSheet(this.sheets[0].name);
  }

  private buildTabs(): void {
    if (!this.tabsEl) return;
    this.tabsEl.empty();
    for (const { name } of this.sheets) {
      const tab = this.tabsEl.createEl("button", {
        text: name,
        cls: "docx-claude-xlsx-tab",
        attr: { title: name, "aria-label": `Switch to sheet ${name}` },
      });
      tab.dataset.sheet = name;
      tab.addEventListener("click", (ev) => {
        ev.preventDefault();
        this.showSheet(name);
      });
    }
  }

  private showSheet(name: string): void {
    if (!this.gridEl) return;
    const entry = this.sheets.find((s) => s.name === name);
    if (!entry) return;
    this.gridEl.empty();
    renderSheetIntoGrid(entry.worksheet, this.gridEl);
    this.updateActiveTab(name);
  }

  private updateActiveTab(active: string): void {
    if (!this.tabsEl) return;
    const tabs = this.tabsEl.querySelectorAll<HTMLElement>(".docx-claude-xlsx-tab");
    tabs.forEach((t) => t.toggleClass("is-active", t.dataset.sheet === active));
  }
}

// ---------------------------------------------------------------------------
// Sheet → HTML grid renderer
// ---------------------------------------------------------------------------

// Excel column width units are "characters of the default font, rendered with
// the maximum digit width." 7px per char is a reasonable approximation for
// Calibri 11 on a typical Mac.
const COL_WIDTH_PX_PER_CH = 7;
const DEFAULT_COL_WIDTH_CH = 8.43;
const POINTS_TO_PX = 4 / 3;

interface MergeRect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

function renderSheetIntoGrid(ws: ExcelJS.Worksheet, container: HTMLElement): void {
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

  const skipMap = computeMergeSkipMap(merges);
  const mergeAnchor = new Map<string, MergeRect>();
  for (const m of merges) mergeAnchor.set(`${m.top}:${m.left}`, m);

  const table = container.createEl("table", { cls: "docx-claude-xlsx-table" });

  const colgroup = table.createEl("colgroup");
  colgroup.createEl("col", { cls: "docx-claude-xlsx-rowhdr-col" });
  for (let c = 1; c <= lastCol; c++) {
    const col = ws.getColumn(c);
    const widthCh = col.width ?? DEFAULT_COL_WIDTH_CH;
    const widthPx = Math.round(widthCh * COL_WIDTH_PX_PER_CH);
    const colEl = colgroup.createEl("col");
    colEl.setAttribute("style", `width: ${widthPx}px`);
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
  for (let r = 1; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const tr = tbody.createEl("tr");
    if (row.height) {
      tr.setAttribute("style", `height: ${Math.round(row.height * POINTS_TO_PX)}px`);
    }
    tr.createEl("th", {
      cls: "docx-claude-xlsx-rowhdr",
      text: String(r),
    });
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
      const text = cellText(cell);
      if (text) td.setText(text);
      const style = cellInlineStyle(cell);
      if (style) td.setAttribute("style", style);
    }
  }
}

function collectMerges(ws: ExcelJS.Worksheet): MergeRect[] {
  const merges: MergeRect[] = [];
  const list = (ws as unknown as { model?: { merges?: string[] } }).model?.merges;
  if (!Array.isArray(list)) return merges;
  for (const addr of list) {
    const m = parseMergeRange(addr);
    if (m) merges.push(m);
  }
  return merges;
}

function parseMergeRange(addr: string): MergeRect | null {
  const m = addr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return null;
  return {
    top: parseInt(m[2], 10),
    left: colNum(m[1]),
    bottom: parseInt(m[4], 10),
    right: colNum(m[3]),
  };
}

function computeMergeSkipMap(merges: MergeRect[]): Set<string> {
  const skip = new Set<string>();
  for (const m of merges) {
    for (let r = m.top; r <= m.bottom; r++) {
      for (let c = m.left; c <= m.right; c++) {
        if (r === m.top && c === m.left) continue;
        skip.add(`${r}:${c}`);
      }
    }
  }
  return skip;
}

function colLetter(c: number): string {
  let s = "";
  while (c > 0) {
    const rem = (c - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function colNum(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function cellText(cell: ExcelJS.Cell): string {
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
      } catch {
        // fall through
      }
    }
    return raw ? "TRUE" : "FALSE";
  }
  if (raw instanceof Date) {
    if (fmt && fmt !== "General") {
      try {
        return numfmtFormat(fmt, dateToSerial(raw));
      } catch {
        // fall through to default
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
  } catch {
    return formatGeneralNumber(raw);
  }
}

function extractCellValue(
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

function formatGeneralNumber(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  if (Number.isInteger(n)) return String(n);
  // Trim FP artifacts (e.g. 22045.858999999997 → 22045.859). Excel's General
  // format shows up to ~11 significant digits, so ~10 decimals is plenty.
  const rounded = Math.round(n * 1e10) / 1e10;
  return String(rounded);
}

function cellInlineStyle(cell: ExcelJS.Cell): string {
  const parts: string[] = [];

  const font = cell.font as Partial<ExcelJS.Font> | undefined;
  if (font) {
    if (font.bold) parts.push("font-weight: 600");
    if (font.italic) parts.push("font-style: italic");
    if (font.underline) parts.push("text-decoration: underline");
    if (typeof font.size === "number") parts.push(`font-size: ${font.size}pt`);
    if (font.name) parts.push(`font-family: ${quoteFont(font.name)}`);
    const colorCss = resolveExcelColor(font.color as ExcelColorRef | undefined);
    if (colorCss) parts.push(`color: ${colorCss}`);
  }

  const fill = cell.fill as ExcelJS.Fill | undefined;
  if (fill && fill.type === "pattern") {
    // For solid patterns fgColor is the fill. For non-solid patterns we use
    // fgColor as the primary color (good enough for a viewer) and fall back
    // to bgColor if fgColor is missing.
    const fgRef = (fill as { fgColor?: ExcelColorRef }).fgColor;
    const bgRef = (fill as { bgColor?: ExcelColorRef }).bgColor;
    const css = resolveExcelColor(fgRef) ?? resolveExcelColor(bgRef);
    if (css) parts.push(`background-color: ${css}`);
  }

  const border = cell.border as Partial<ExcelJS.Borders> | undefined;
  if (border) {
    if (border.top?.style) parts.push(`border-top: ${borderCss(border.top)}`);
    if (border.bottom?.style) parts.push(`border-bottom: ${borderCss(border.bottom)}`);
    if (border.left?.style) parts.push(`border-left: ${borderCss(border.left)}`);
    if (border.right?.style) parts.push(`border-right: ${borderCss(border.right)}`);
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

function argbToCss(argb: string): string {
  // ExcelJS hands us 8 hex chars (alpha first). CSS hex with alpha is RRGGBBAA.
  // Pad short inputs to a full 8-hex ARGB so we never emit malformed CSS.
  const padded = argb.length >= 8 ? argb : argb.padStart(8, "F");
  const a = padded.slice(-8, -6).toUpperCase();
  const rgb = padded.slice(-6);
  return a === "FF" ? `#${rgb}` : `#${rgb}${a}`;
}

function borderCss(b: Partial<ExcelJS.Border>): string {
  const style = b.style ?? "thin";
  const width = style === "thick" || style === "double" ? "2px" : "1px";
  const css = style === "double" ? "double" : "solid";
  const colorCss = resolveExcelColor(b.color as ExcelColorRef | undefined) ?? "#d4d4d4";
  return `${width} ${css} ${colorCss}`;
}

// ---------------------------------------------------------------------------
// Color resolution: argb | theme + tint | indexed → CSS color
// ---------------------------------------------------------------------------

interface ExcelColorRef {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

// Default Office theme color palette (OOXML clrScheme). The vast majority of
// workbooks use this palette; for workbooks with a custom theme the colors
// will be slightly off but the tint math still produces a sensible variant.
const DEFAULT_THEME_RGB = [
  "FFFFFF", // 0: lt1 (Background 1)
  "000000", // 1: dk1 (Text 1)
  "E7E6E6", // 2: lt2 (Background 2)
  "44546A", // 3: dk2 (Text 2)
  "4472C4", // 4: Accent 1 (blue)
  "ED7D31", // 5: Accent 2 (orange)
  "A5A5A5", // 6: Accent 3 (gray)
  "FFC000", // 7: Accent 4 (gold)
  "5B9BD5", // 8: Accent 5 (light blue)
  "70AD47", // 9: Accent 6 (green)
  "0563C1", // 10: Hyperlink
  "954F72", // 11: Followed Hyperlink
];

// Standard Excel indexed color palette (xlsx legacy "indexed" attribute).
const INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

function resolveExcelColor(color: ExcelColorRef | undefined | null): string | null {
  if (!color) return null;
  if (color.argb) return argbToCss(color.argb);
  if (typeof color.theme === "number") {
    const hex = DEFAULT_THEME_RGB[color.theme];
    if (hex) return applyTint(hex, color.tint ?? 0);
    // Fall through: workbooks can carry both a theme ref and an indexed
    // fallback. If the theme index is out of range (custom themes can
    // exceed 11), try indexed before giving up.
  }
  if (typeof color.indexed === "number") {
    const hex = INDEXED_COLORS[color.indexed];
    if (hex) return `#${hex}`;
  }
  return null;
}

// OOXML tint is applied in HSL space against luminance. Positive tint lightens
// toward white, negative tint darkens toward black.
function applyTint(hex6: string, tint: number): string {
  if (Math.abs(tint) < 0.001) return `#${hex6}`;
  const r = parseInt(hex6.slice(0, 2), 16) / 255;
  const g = parseInt(hex6.slice(2, 4), 16) / 255;
  const b = parseInt(hex6.slice(4, 6), 16) / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0;
  let s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }

  const newL = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;

  let nr: number;
  let ng: number;
  let nb: number;
  if (s === 0) {
    nr = ng = nb = newL;
  } else {
    const q = newL < 0.5 ? newL * (1 + s) : newL + s - newL * s;
    const p = 2 * newL - q;
    const hueToRgb = (t: number): number => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    nr = hueToRgb(h + 1 / 3);
    ng = hueToRgb(h);
    nb = hueToRgb(h - 1 / 3);
  }

  const toHex = (n: number): string => {
    const v = Math.round(Math.max(0, Math.min(255, n * 255)));
    return v.toString(16).padStart(2, "0").toUpperCase();
  };
  return `#${toHex(nr)}${toHex(ng)}${toHex(nb)}`;
}
