import { NS, directChild, childrenNS, elementChildren } from "./ooxml";
import { EMU_PER_PT } from "./geometry";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import { resolveThemeFont, parseBulletFrom, type RunStyleDefaults, type BulletDef } from "./inheritance";
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
  marLEmu: number | null; // paragraph left margin
  indentEmu: number | null; // first-line indent (negative = hanging)
  // resolved bullet (the "none"/unknown cases collapse to null)
  bullet: Exclude<BulletDef, { kind: "none" }> | null;
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

    const marLAttr = pPr?.getAttribute("marL");
    const marLEmu = marLAttr != null ? parseInt(marLAttr, 10) : defs.marL ?? null;
    const indentAttr = pPr?.getAttribute("indent");
    const indentEmu = indentAttr != null ? parseInt(indentAttr, 10) : defs.indent ?? null;
    // Explicit pPr bullet wins; otherwise inherit the list style's bullet. A
    // resolved {kind:"none"} (or unknown) renders no marker.
    const buDef = parseBulletFrom(pPr) ?? defs.bullet;
    const bullet = buDef && buDef.kind !== "none" ? buDef : null;

    const runs: TextRun[] = [];
    for (const child of elementChildren(p)) {
      if (child.namespaceURI !== NS.a) continue;
      if (child.localName === "r" || child.localName === "fld") {
        runs.push(parseRun(child, ctx, defs));
      } else if (child.localName === "br") {
        runs.push(blankRun(true));
      }
    }
    out.push({ runs, align, level, marLEmu, indentEmu, bullet });
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

// PowerPoint's default body indents when a level omits marL/indent: ~0.375" of
// left margin per level with a matching hanging indent for the bullet.
const DEFAULT_BULLET_MARL = 342900; // EMU (0.375")
const DEFAULT_BULLET_INDENT = -342900;

// buChar glyphs are usually private-use codepoints that only render in their
// buFont (Symbol/Wingdings), which may not be installed. Map the common ones to
// portable Unicode so a bullet always shows instead of tofu; fall back to a
// round bullet for any other private-use char, else render the char as given.
const BULLET_GLYPHS: Record<number, string> = {
  0x2022: "•", 0x00b7: "•", 0xf0b7: "•", // Symbol bullet / middle dot
  0xf06c: "●", // Wingdings 'l' -> black circle
  0xf0a7: "▪", 0xf0a8: "▪", // small black square
  0xf0d8: "➢", // arrowhead
  0xf0fc: "✔", // check mark
  0xf02d: "–", // Symbol hyphen -> en dash
};

function bulletGlyph(char: string, font: string | null): { text: string; fontFamily?: string } {
  const cp = char.codePointAt(0) ?? 0;
  if (BULLET_GLYPHS[cp]) return { text: BULLET_GLYPHS[cp] };
  if (cp >= 0xf000 && cp <= 0xf0ff) return { text: "•" }; // unknown PUA -> bullet
  return { text: char, fontFamily: font ? `"${font}"` : undefined };
}

// Render parsed paragraphs into `el`. `scale` is px-per-EMU for this slide, so a
// point size maps to px as sizePt * EMU_PER_PT * scale.
export function renderParagraphsInto(
  paras: Paragraph[],
  el: HTMLElement,
  scale: number,
  fontScale = 1,
): void {
  const counters: number[] = []; // auto-numbering state, indexed by level

  for (const para of paras) {
    const pEl = el.createDiv({ cls: "docx-claude-pptx-p" });
    if (para.align) pEl.style.textAlign = mapAlign(para.align);

    const hasBullet = !!para.bullet && para.runs.length > 0;
    const marLEmu = para.marLEmu ?? (para.bullet ? DEFAULT_BULLET_MARL * (para.level + 1) : null);
    const indentEmu = para.indentEmu ?? (para.bullet ? DEFAULT_BULLET_INDENT : null);
    const marLpx = marLEmu != null ? marLEmu * scale : null;
    const hangPx = indentEmu != null ? Math.max(0, -indentEmu * scale) : 0;

    // Sequence auto-numbers per level; a deeper level or a break resets counts.
    if (para.bullet?.kind === "autonum") {
      counters[para.level] = (counters[para.level] ?? 0) + 1;
      counters.length = para.level + 1;
    } else if (!para.bullet) {
      counters.length = Math.min(counters.length, para.level);
    }

    const emitRuns = (host: HTMLElement): void => {
      if (para.runs.length === 0) {
        host.createEl("br");
        return;
      }
      for (const run of para.runs) {
        if (run.lineBreak) {
          host.createEl("br");
          continue;
        }
        if (run.text === "") continue;
        const span = host.createEl("span", { text: run.text });
        if (run.bold) span.style.fontWeight = "bold";
        if (run.italic) span.style.fontStyle = "italic";
        if (run.underline) span.style.textDecoration = "underline";
        const pt = (run.sizePt ?? DEFAULT_FONT_PT) * fontScale;
        span.style.fontSize = `${pt * EMU_PER_PT * scale}px`;
        if (run.colorCss) span.style.color = run.colorCss;
        if (run.fontFamily) span.style.fontFamily = run.fontFamily;
      }
    };

    if (hasBullet && para.bullet) {
      // Hanging-indent layout: bullet sits in a fixed gutter, text flows after.
      pEl.style.display = "flex";
      pEl.style.alignItems = "baseline";
      if (marLpx != null) pEl.style.paddingInlineStart = `${Math.max(0, marLpx - hangPx)}px`;
      const firstPt =
        (para.runs.find((r) => !r.lineBreak && r.text)?.sizePt ?? DEFAULT_FONT_PT) * fontScale;
      const buEl = pEl.createEl("span", { cls: "docx-claude-pptx-bullet" });
      if (para.bullet.kind === "autonum") {
        buEl.setText(`${counters[para.level] ?? 1}.`);
      } else {
        const g = bulletGlyph(para.bullet.char, para.bullet.font);
        buEl.setText(g.text);
        if (g.fontFamily) buEl.style.fontFamily = g.fontFamily;
      }
      buEl.style.flex = `0 0 ${hangPx || 16}px`;
      buEl.style.fontSize = `${firstPt * EMU_PER_PT * scale}px`;
      const txt = pEl.createEl("span");
      txt.style.flex = "1";
      emitRuns(txt);
    } else {
      if (marLpx != null) pEl.style.paddingInlineStart = `${marLpx}px`;
      else if (para.level > 0) pEl.style.marginInlineStart = `${para.level * 1.2}em`;
      emitRuns(pEl);
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
