import type ExcelJS from "exceljs";
import { ExcelColorRef, resolveExcelColor } from "./colors";

export interface RichTextRun {
  text?: string;
  font?: Partial<ExcelJS.Font>;
}

export function applyRunFontToElement(
  font: Partial<ExcelJS.Font> | undefined,
  el: HTMLElement,
  theme?: readonly string[],
): void {
  if (!font) return;
  const s = el.style;
  // Excel "bold" is weight 700, not 600.
  if (font.bold) s.fontWeight = "bold";
  if (font.italic) s.fontStyle = "italic";
  // strike and underline share text-decoration; merge so one doesn't clobber
  // the other. "double" underline keeps the double style.
  const decorations: string[] = [];
  if (font.underline) {
    decorations.push(font.underline === "double" ? "underline double" : "underline");
  }
  if (font.strike) decorations.push("line-through");
  if (decorations.length) s.textDecoration = decorations.join(" ");
  if (typeof font.size === "number") s.fontSize = `${font.size}pt`;
  if (font.name) s.fontFamily = quoteFont(font.name);
  if (font.vertAlign === "superscript") s.verticalAlign = "super";
  else if (font.vertAlign === "subscript") s.verticalAlign = "sub";
  const colorCss = resolveExcelColor(font.color as ExcelColorRef | undefined, theme);
  if (colorCss) s.color = colorCss;
}

export function renderRichTextRuns(
  runs: RichTextRun[],
  target: HTMLElement,
  theme?: readonly string[],
): void {
  for (const run of runs) {
    if (!run.text) continue;
    const span = target.createSpan();
    span.setText(run.text);
    applyRunFontToElement(run.font, span, theme);
  }
}

function quoteFont(name: string): string {
  return /^[A-Za-z0-9_-]+$/.test(name) ? name : `"${name.replace(/"/g, "")}"`;
}
