import type ExcelJS from "exceljs";
import { ExcelColorRef, resolveExcelColor } from "../colors";

// Per-cell set of style-property groups a higher-precedence CF rule has already
// written, keyed by "r:c". Rules are applied highest-precedence first, so once a
// group is locked a lower-precedence rule must not overwrite it. "background"
// covers solid fill, color-scale fill, and data-bar gradient (they all paint the
// cell background and would clobber each other). Non-conflicting groups still
// stack (e.g. bold from one rule + red fill from another).
export type CfLocks = Map<string, Set<string>>;

export function lockSetFor(locks: CfLocks, key: string): Set<string> {
  let s = locks.get(key);
  if (!s) {
    s = new Set();
    locks.set(key, s);
  }
  return s;
}

// Apply a conditional-formatting rule's dxf style to a rendered cell. CF styles
// most commonly carry font color + fill; we also honor bold/italic/underline.
// `locked` (when passed) enforces per-property precedence: each group is written
// only if not already locked, then locked.
export function mergeStyleIntoElement(
  td: HTMLTableCellElement,
  style: Partial<ExcelJS.Style> | undefined,
  theme?: readonly string[],
  locked?: Set<string>,
): void {
  if (!style) return;
  const write = (group: string, set: () => void): void => {
    if (locked?.has(group)) return;
    set();
    locked?.add(group);
  };

  const font = style.font as Partial<ExcelJS.Font> | undefined;
  if (font) {
    if (font.bold) write("fontWeight", () => (td.style.fontWeight = "bold"));
    if (font.italic) write("fontStyle", () => (td.style.fontStyle = "italic"));
    if (font.underline) write("textDecoration", () => (td.style.textDecoration = "underline"));
    const c = resolveExcelColor(font.color as ExcelColorRef | undefined, theme);
    if (c) write("color", () => (td.style.color = c));
  }

  const fill = style.fill as ExcelJS.Fill | undefined;
  if (fill && fill.type === "pattern") {
    const fg = (fill as { fgColor?: ExcelColorRef }).fgColor;
    const bg = (fill as { bgColor?: ExcelColorRef }).bgColor;
    const c = resolveExcelColor(fg, theme) ?? resolveExcelColor(bg, theme);
    if (c) write("background", () => (td.style.backgroundColor = c));
  }
}
