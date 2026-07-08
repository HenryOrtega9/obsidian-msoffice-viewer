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
    // Merge with any decoration the base cell style already set (a bare
    // assignment clobbers an existing line-through/underline).
    if (font.underline || font.strike) {
      write("textDecoration", () => {
        const existing = td.style.textDecoration
          .split(/\s+/)
          .filter((t) => t && t !== "none");
        if (font.underline && !existing.includes("underline")) existing.push("underline");
        if (font.strike && !existing.includes("line-through")) existing.push("line-through");
        td.style.textDecoration = existing.join(" ");
      });
    }
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

  const border = style.border as Partial<ExcelJS.Borders> | undefined;
  if (border) {
    const sideCss = (b: Partial<ExcelJS.Border> | undefined): string | null => {
      if (!b?.style) return null;
      const color = resolveExcelColor(b.color as ExcelColorRef | undefined, theme) ?? "#000000";
      const width = b.style === "medium" ? "2px" : b.style === "thick" ? "3px" : "1px";
      return `${width} solid ${color}`;
    };
    const sides: ["top" | "bottom" | "left" | "right", string][] = [
      ["top", "borderTop"], ["bottom", "borderBottom"], ["left", "borderLeft"], ["right", "borderRight"],
    ];
    for (const [side, prop] of sides) {
      const css = sideCss(border[side]);
      if (css) {
        write(`border-${side}`, () => {
          (td.style as unknown as Record<string, string>)[prop] = css;
        });
      }
    }
  }
}
