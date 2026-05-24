import { NS, directChild, childrenNS, firstChildNS } from "./ooxml";
import { findSpTree } from "./presentation";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import type { PptxTheme, ClrMap } from "./themes";

export interface Ph {
  type: string;
  idx: string;
}

export function getPh(sp: Element): Ph | null {
  const nvSpPr = directChild(sp, NS.p, "nvSpPr");
  const nvPr = nvSpPr ? directChild(nvSpPr, NS.p, "nvPr") : null;
  const ph = nvPr ? directChild(nvPr, NS.p, "ph") : null;
  if (!ph) return null;
  return { type: ph.getAttribute("type") ?? "body", idx: ph.getAttribute("idx") ?? "" };
}

// Resolve a placeholder shape's geometry through the inheritance chain:
// layout first (most specific override of the master), then master. Returns the
// matching a:xfrm element, or null if none defines geometry.
export function resolvePlaceholderXfrm(
  sp: Element,
  layoutDoc: Document | null,
  masterDoc: Document | null,
): Element | null {
  const ph = getPh(sp);
  if (!ph) return null;
  for (const doc of [layoutDoc, masterDoc]) {
    const match = findPhShape(doc, ph);
    if (!match) continue;
    const spPr = directChild(match, NS.p, "spPr");
    const xfrm = spPr ? directChild(spPr, NS.a, "xfrm") : null;
    if (xfrm) return xfrm;
  }
  return null;
}

function findPhShape(doc: Document | null, ph: Ph): Element | null {
  if (!doc) return null;
  const spTree = findSpTree(doc);
  if (!spTree) return null;
  const shapes = childrenNS(spTree, NS.p, "sp");
  const exact = shapes.find((s) => {
    const p = getPh(s);
    return p && p.type === ph.type && p.idx === ph.idx;
  });
  return exact ?? shapes.find((s) => {
    const p = getPh(s);
    return p && p.type === ph.type;
  }) ?? null;
}

export interface RunStyleDefaults {
  sizePt?: number;
  colorCss?: string | null;
  bold?: boolean;
  fontFamily?: string | null;
  align?: string | null;
}

// Build a per-level resolver of default run styling from the master's
// p:txStyles, picked by placeholder type. Covers the common case where a slide
// omits run sizes and relies on the template (e.g. titles at 44pt).
export function buildPhDefaults(
  phType: string | null,
  masterDoc: Document | null,
  theme: PptxTheme | null,
  clrMap: ClrMap,
): (level: number) => RunStyleDefaults {
  const styleEl = pickStyleEl(phType, masterDoc);
  return (level: number): RunStyleDefaults => {
    if (!styleEl) return {};
    const lvlPr = nthLvlPr(styleEl, level);
    if (!lvlPr) return {};
    const out: RunStyleDefaults = { align: lvlPr.getAttribute("algn") };
    const defRPr = directChild(lvlPr, NS.a, "defRPr");
    if (defRPr) {
      const sz = defRPr.getAttribute("sz");
      if (sz) {
        const n = parseInt(sz, 10);
        if (!Number.isNaN(n)) out.sizePt = n / 100;
      }
      const b = defRPr.getAttribute("b");
      if (b != null) out.bold = b === "1";
      const fill = directChild(defRPr, NS.a, "solidFill");
      if (fill) out.colorCss = resolveDrawingMlColor(parseColorChoice(fill), theme, clrMap);
      const latin = directChild(defRPr, NS.a, "latin");
      if (latin) out.fontFamily = resolveThemeFont(latin.getAttribute("typeface"), theme);
    }
    return out;
  };
}

function pickStyleEl(phType: string | null, masterDoc: Document | null): Element | null {
  if (!masterDoc) return null;
  const txStyles = firstChildNS(masterDoc.documentElement, NS.p, "txStyles");
  if (!txStyles) return null;
  const which =
    phType === "title" || phType === "ctrTitle"
      ? "titleStyle"
      : phType === "dt" || phType === "ftr" || phType === "sldNum"
        ? "otherStyle"
        : "bodyStyle";
  return directChild(txStyles, NS.p, which);
}

function nthLvlPr(styleEl: Element, level: number): Element | null {
  const name = `lvl${Math.min(8, Math.max(0, level)) + 1}pPr`;
  return directChild(styleEl, NS.a, name);
}

// Map a typeface to a CSS font stack, resolving the +mj-lt / +mn-lt theme-font
// references to the theme's major/minor font.
export function resolveThemeFont(typeface: string | null, theme: PptxTheme | null): string | null {
  if (!typeface) return null;
  let name = typeface;
  if (/^\+mj/i.test(typeface)) name = theme?.majorFont ?? "";
  else if (/^\+mn/i.test(typeface)) name = theme?.minorFont ?? "";
  name = name.replace(/^\+[a-z]+-/i, "").replace(/"/g, "").trim();
  if (!name) return null;
  return `"${name}", sans-serif`;
}
