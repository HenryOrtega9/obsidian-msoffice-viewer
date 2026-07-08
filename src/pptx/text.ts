import { NS, directChild, childrenNS, elementChildren } from "./ooxml";
import { EMU_PER_PT } from "./geometry";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import {
  resolveThemeFont,
  parseBulletFrom,
  readLvlPrDefaults,
  type RunStyleDefaults,
  type BulletDef,
} from "./inheritance";
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
  strike: boolean;
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
  lnSpcPct: number | null; // line spacing as fraction of single spacing
  lnSpcPts: number | null; // exact line spacing in points
  spcBefPts: number | null;
  spcAftPts: number | null;
}

const DEFAULT_FONT_PT = 18;

// Per-level resolver of inherited run defaults (from the shape/layout/master
// list-style chain). Optional: table cells pass none and fall back to literals.
export type DefaultsFor = (level: number) => RunStyleDefaults;

export function parseTxBody(txBody: Element, ctx: TextResolveCtx, defaultsFor?: DefaultsFor): Paragraph[] {
  const out: Paragraph[] = [];
  for (const p of childrenNS(txBody, NS.a, "p")) {
    const pPr = directChild(p, NS.a, "pPr");
    const level = pPr ? parseInt(pPr.getAttribute("lvl") ?? "0", 10) || 0 : 0;
    const inherited = defaultsFor ? defaultsFor(level) : {};
    // The paragraph's own pPr (incl. its defRPr) overrides the inherited chain.
    const own = pPr ? readLvlPrDefaults(pPr, ctx.theme, ctx.clrMap) : {};
    const defs: RunStyleDefaults = { ...inherited };
    for (const [k, v] of Object.entries(own)) {
      if (v !== undefined) (defs as Record<string, unknown>)[k] = v;
    }

    const align = defs.align ?? null;
    const marLEmu = defs.marL ?? null;
    const indentEmu = defs.indent ?? null;
    // Explicit pPr bullet wins; otherwise inherit the list style's bullet. A
    // resolved {kind:"none"} (or unknown) renders no marker.
    const buDef = parseBulletFrom(pPr) ?? inherited.bullet;
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
    out.push({
      runs,
      align,
      level,
      marLEmu,
      indentEmu,
      bullet,
      lnSpcPct: defs.lnSpcPct ?? null,
      lnSpcPts: defs.lnSpcPts ?? null,
      spcBefPts: defs.spcBefPts ?? null,
      spcAftPts: defs.spcAftPts ?? null,
    });
  }
  return out;
}

function parseRun(r: Element, ctx: TextResolveCtx, defs: RunStyleDefaults): TextRun {
  const t = directChild(r, NS.a, "t");
  const run = blankRun(false);
  run.text = t?.textContent ?? "";

  // Seed from inherited defaults, then let explicit rPr override.
  run.bold = defs.bold ?? false;
  run.italic = defs.italic ?? false;
  run.sizePt = defs.sizePt ?? null;
  run.colorCss = defs.colorCss ?? null;
  run.fontFamily = defs.fontFamily ?? null;

  const rPr = directChild(r, NS.a, "rPr");
  if (rPr) {
    const b = rPr.getAttribute("b");
    if (b != null) run.bold = b === "1";
    const i = rPr.getAttribute("i");
    if (i != null) run.italic = i === "1";
    const u = rPr.getAttribute("u");
    run.underline = !!u && u !== "none";
    const strike = rPr.getAttribute("strike");
    run.strike = !!strike && strike !== "noStrike";
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
    // Hyperlink runs take the theme hlink color and underline unless the run
    // explicitly overrides them.
    if (directChild(rPr, NS.a, "hlinkClick")) {
      if (!fill) {
        const hlink = resolveDrawingMlColor({ scheme: "hlink", mods: [] }, ctx.theme, ctx.clrMap);
        if (hlink) run.colorCss = hlink;
      }
      if (u == null) run.underline = true;
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
    strike: false,
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

const ROMAN: [number, string][] = [
  [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
  [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
];

function toRoman(n: number): string {
  let out = "";
  let rest = Math.max(1, n);
  for (const [v, s] of ROMAN) {
    while (rest >= v) {
      out += s;
      rest -= v;
    }
  }
  return out;
}

function toAlpha(n: number): string {
  let out = "";
  let rest = Math.max(1, n);
  while (rest > 0) {
    rest -= 1;
    out = String.fromCharCode(97 + (rest % 26)) + out;
    rest = Math.floor(rest / 26);
  }
  return out;
}

// Format an auto-number per its ST_TextAutonumberScheme: a base numeral style
// (arabic / alpha / roman, upper or lower) plus a punctuation wrapper
// (Period / ParenR / ParenBoth / plain).
export function formatAutoNum(type: string, n: number): string {
  let core: string;
  if (/^alphaUc/.test(type)) core = toAlpha(n).toUpperCase();
  else if (/^alphaLc/.test(type)) core = toAlpha(n);
  else if (/^romanUc/.test(type)) core = toRoman(n).toUpperCase();
  else if (/^romanLc/.test(type)) core = toRoman(n);
  else core = String(n);
  if (type.endsWith("ParenBoth")) return `(${core})`;
  if (type.endsWith("ParenR")) return `${core})`;
  if (type.endsWith("Period")) return `${core}.`;
  return core;
}

// Render parsed paragraphs into `el`. `scale` is px-per-EMU for this slide, so a
// point size maps to px as sizePt * EMU_PER_PT * scale. `fontScale` and
// `lnSpcReduction` come from normAutofit.
export function renderParagraphsInto(
  paras: Paragraph[],
  el: HTMLElement,
  scale: number,
  fontScale = 1,
  lnSpcReduction = 0,
): void {
  const counters: number[] = []; // auto-numbering state, indexed by level

  for (const para of paras) {
    const pEl = el.createDiv({ cls: "docx-claude-pptx-p" });
    if (para.align) pEl.style.textAlign = mapAlign(para.align);

    // Line spacing: spcPct is a fraction of single spacing (Office single is
    // ~1.2), spcPts an exact point height. Autofit's lnSpcReduction shrinks it.
    const spcFactor = 1 - lnSpcReduction;
    if (para.lnSpcPts != null) {
      pEl.style.lineHeight = `${para.lnSpcPts * EMU_PER_PT * scale * spcFactor}px`;
    } else if (para.lnSpcPct != null) {
      pEl.style.lineHeight = `${1.2 * para.lnSpcPct * spcFactor}`;
    } else if (lnSpcReduction > 0) {
      pEl.style.lineHeight = `${1.2 * spcFactor}`;
    }
    if (para.spcBefPts != null) pEl.style.marginTop = `${para.spcBefPts * EMU_PER_PT * scale}px`;
    if (para.spcAftPts != null) pEl.style.marginBottom = `${para.spcAftPts * EMU_PER_PT * scale}px`;

    const hasBullet = !!para.bullet && para.runs.length > 0;
    const marLEmu = para.marLEmu ?? (para.bullet ? DEFAULT_BULLET_MARL * (para.level + 1) : null);
    const indentEmu = para.indentEmu ?? (para.bullet ? DEFAULT_BULLET_INDENT : null);
    const marLpx = marLEmu != null ? marLEmu * scale : null;
    const hangPx = indentEmu != null ? Math.max(0, -indentEmu * scale) : 0;

    // Sequence auto-numbers per level; a deeper level or a break resets counts.
    if (para.bullet?.kind === "autonum") {
      counters[para.level] = counters[para.level] ?? para.bullet.startAt - 1;
      counters[para.level] += 1;
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
        const deco = [
          run.underline ? "underline" : "",
          run.strike ? "line-through" : "",
        ].filter(Boolean).join(" ");
        if (deco) span.style.textDecoration = deco;
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
        buEl.setText(formatAutoNum(para.bullet.type, counters[para.level] ?? 1));
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
