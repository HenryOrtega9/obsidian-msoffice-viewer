import { NS, directChild, childrenNS, elementChildren } from "./ooxml";
import { EMU_PER_PT } from "./geometry";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import { resolveThemeFont, type RunStyleDefaults } from "./inheritance";
import type { PptxTheme, ClrMap } from "./themes";

export interface TextResolveCtx {
  theme: PptxTheme | null;
  clrMap: ClrMap;
}

export interface TextRun {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  sizePt: number | null;
  colorCss: string | null;
  fontFamily: string | null; // CSS-ready font stack
  lineBreak?: boolean;
}

export interface Paragraph {
  runs: TextRun[];
  align: string | null; // OOXML algn: l/ctr/r/just
  level: number;
}

const DEFAULT_FONT_PT = 18;

// Per-level resolver of inherited run defaults (from master txStyles). Optional:
// non-placeholder shapes and table cells pass none and fall back to literals.
export type DefaultsFor = (level: number) => RunStyleDefaults;

export function parseTxBody(txBody: Element, ctx: TextResolveCtx, defaultsFor?: DefaultsFor): Paragraph[] {
  const out: Paragraph[] = [];
  for (const p of childrenNS(txBody, NS.a, "p")) {
    const pPr = directChild(p, NS.a, "pPr");
    const level = pPr ? parseInt(pPr.getAttribute("lvl") ?? "0", 10) || 0 : 0;
    const defs = defaultsFor ? defaultsFor(level) : {};
    const align = pPr?.getAttribute("algn") ?? defs.align ?? null;
    const runs: TextRun[] = [];
    for (const child of elementChildren(p)) {
      if (child.namespaceURI !== NS.a) continue;
      if (child.localName === "r" || child.localName === "fld") {
        runs.push(parseRun(child, ctx, defs));
      } else if (child.localName === "br") {
        runs.push(blankRun(true));
      }
    }
    out.push({ runs, align, level });
  }
  return out;
}

function parseRun(r: Element, ctx: TextResolveCtx, defs: RunStyleDefaults): TextRun {
  const t = directChild(r, NS.a, "t");
  const run = blankRun(false);
  run.text = t?.textContent ?? "";

  // Seed from inherited defaults, then let explicit rPr override.
  run.bold = defs.bold ?? false;
  run.sizePt = defs.sizePt ?? null;
  run.colorCss = defs.colorCss ?? null;
  run.fontFamily = defs.fontFamily ?? null;

  const rPr = directChild(r, NS.a, "rPr");
  if (rPr) {
    const b = rPr.getAttribute("b");
    if (b != null) run.bold = b === "1";
    run.italic = rPr.getAttribute("i") === "1";
    const u = rPr.getAttribute("u");
    run.underline = !!u && u !== "none";
    const sz = rPr.getAttribute("sz");
    if (sz) {
      const n = parseInt(sz, 10);
      if (!Number.isNaN(n)) run.sizePt = n / 100;
    }
    const fill = directChild(rPr, NS.a, "solidFill");
    if (fill) {
      const css = resolveDrawingMlColor(parseColorChoice(fill), ctx.theme, ctx.clrMap);
      if (css) run.colorCss = css;
    }
    const latin = directChild(rPr, NS.a, "latin");
    if (latin) {
      const css = resolveThemeFont(latin.getAttribute("typeface"), ctx.theme);
      if (css) run.fontFamily = css;
    }
  }
  return run;
}

function blankRun(lineBreak: boolean): TextRun {
  return {
    text: "",
    bold: false,
    italic: false,
    underline: false,
    sizePt: null,
    colorCss: null,
    fontFamily: null,
    lineBreak,
  };
}

// Render parsed paragraphs into `el`. `scale` is px-per-EMU for this slide, so a
// point size maps to px as sizePt * EMU_PER_PT * scale.
export function renderParagraphsInto(paras: Paragraph[], el: HTMLElement, scale: number): void {
  for (const para of paras) {
    const pEl = el.createDiv({ cls: "docx-claude-pptx-p" });
    if (para.align) pEl.style.textAlign = mapAlign(para.align);
    if (para.level > 0) pEl.style.marginInlineStart = `${para.level * 1.2}em`;

    if (para.runs.length === 0) {
      pEl.createEl("br");
      continue;
    }
    for (const run of para.runs) {
      if (run.lineBreak) {
        pEl.createEl("br");
        continue;
      }
      if (run.text === "") continue;
      const span = pEl.createEl("span", { text: run.text });
      if (run.bold) span.style.fontWeight = "bold";
      if (run.italic) span.style.fontStyle = "italic";
      if (run.underline) span.style.textDecoration = "underline";
      const pt = run.sizePt ?? DEFAULT_FONT_PT;
      span.style.fontSize = `${pt * EMU_PER_PT * scale}px`;
      if (run.colorCss) span.style.color = run.colorCss;
      if (run.fontFamily) span.style.fontFamily = run.fontFamily;
    }
  }
}

function mapAlign(algn: string): string {
  switch (algn) {
    case "ctr": return "center";
    case "r": return "right";
    case "just": return "justify";
    default: return "left";
  }
}
