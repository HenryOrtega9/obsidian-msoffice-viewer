import type ExcelJS from "exceljs";
import { ExcelColorRef, resolveExcelColor } from "../colors";

// Apply a conditional-formatting rule's dxf style to a rendered cell. CF styles
// most commonly carry font color + fill; we also honor bold/italic/underline.
export function mergeStyleIntoElement(
  td: HTMLTableCellElement,
  style: Partial<ExcelJS.Style> | undefined,
  theme?: readonly string[],
): void {
  if (!style) return;

  const font = style.font as Partial<ExcelJS.Font> | undefined;
  if (font) {
    if (font.bold) td.style.fontWeight = "600";
    if (font.italic) td.style.fontStyle = "italic";
    if (font.underline) td.style.textDecoration = "underline";
    const c = resolveExcelColor(font.color as ExcelColorRef | undefined, theme);
    if (c) td.style.color = c;
  }

  const fill = style.fill as ExcelJS.Fill | undefined;
  if (fill && fill.type === "pattern") {
    const fg = (fill as { fgColor?: ExcelColorRef }).fgColor;
    const bg = (fill as { bgColor?: ExcelColorRef }).bgColor;
    const c = resolveExcelColor(fg, theme) ?? resolveExcelColor(bg, theme);
    if (c) td.style.backgroundColor = c;
  }
}
