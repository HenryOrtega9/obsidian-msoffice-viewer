import { NS, directChild, firstChildNS } from "./ooxml";
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

const TITLE_FAMILY = new Set(["title", "ctrTitle"]);

function sameTitleFamily(a: string, b: string): boolean {
  return TITLE_FAMILY.has(a) && TITLE_FAMILY.has(b);
}

// Find the layout/master shape a slide placeholder inherits from. Per
// ECMA-376, matching is primarily by idx (a slide ph can re-type the layout's
// placeholder at the same idx); title/ctrTitle are one family matched by type.
// Placeholders can sit inside groups, so search the whole spTree.
export function findPhShape(doc: Document | null, ph: Ph): Element | null {
  if (!doc) return null;
  const spTree = findSpTree(doc);
  if (!spTree) return null;
  const shapes = Array.from(spTree.getElementsByTagNameNS(NS.p, "sp"));
  const phOf = new Map<Element, Ph>();
  const withPh = shapes.filter((s) => {
    const p = getPh(s);
    if (p) phOf.set(s, p);
    return !!p;
  });
  const byPh = (pred: (p: Ph) => boolean): Element | null =>
    withPh.find((s) => pred(phOf.get(s)!)) ?? null;

  return (
    byPh((p) => p.type === ph.type && p.idx === ph.idx) ??
    (ph.idx !== "" ? byPh((p) => p.idx === ph.idx) : null) ??
    (TITLE_FAMILY.has(ph.type) ? byPh((p) => sameTitleFamily(p.type, ph.type)) : null) ??
    byPh((p) => p.type === ph.type)
  );
}

export type BulletDef =
  | { kind: "none" }
  | { kind: "char"; char: string; font: string | null }
  | { kind: "autonum"; type: string; startAt: number };

export interface RunStyleDefaults {
  sizePt?: number;
  colorCss?: string | null;
  bold?: boolean;
  italic?: boolean;
  fontFamily?: string | null;
  align?: string | null;
  marL?: number; // paragraph left margin (EMU)
  indent?: number; // first-line indent (EMU, negative = hanging)
  bullet?: BulletDef;
  lnSpcPct?: number; // line spacing as fraction of single (spcPct/100000)
  lnSpcPts?: number; // exact line spacing in points (spcPts/100)
  spcBefPts?: number;
  spcAftPts?: number;
}

// Read the bullet choice (buNone / buChar / buAutoNum) from a pPr or a master
// lvlNpPr element. Returns undefined when the element specifies no bullet
// choice, so callers can fall back to the inherited list style.
export function parseBulletFrom(el: Element | null | undefined): BulletDef | undefined {
  if (!el) return undefined;
  if (directChild(el, NS.a, "buNone")) return { kind: "none" };
  const buChar = directChild(el, NS.a, "buChar");
  if (buChar) {
    const font = directChild(el, NS.a, "buFont")?.getAttribute("typeface") ?? null;
    return { kind: "char", char: buChar.getAttribute("char") ?? "•", font };
  }
  const buAuto = directChild(el, NS.a, "buAutoNum");
  if (buAuto) {
    const startAt = parseInt(buAuto.getAttribute("startAt") ?? "1", 10) || 1;
    return { kind: "autonum", type: buAuto.getAttribute("type") ?? "arabicPeriod", startAt };
  }
  return undefined;
}

// Extract the defaults a single lvlNpPr (or pPr) element defines. Only sets
// fields the element actually specifies so partial results merge cleanly.
export function readLvlPrDefaults(
  lvlPr: Element,
  theme: PptxTheme | null,
  clrMap: ClrMap,
): RunStyleDefaults {
  const out: RunStyleDefaults = {};
  const algn = lvlPr.getAttribute("algn");
  if (algn) out.align = algn;
  const marL = lvlPr.getAttribute("marL");
  if (marL) out.marL = parseInt(marL, 10);
  const indent = lvlPr.getAttribute("indent");
  if (indent) out.indent = parseInt(indent, 10);
  out.bullet = parseBulletFrom(lvlPr);

  const lnSpc = directChild(lvlPr, NS.a, "lnSpc");
  if (lnSpc) {
    const pct = directChild(lnSpc, NS.a, "spcPct");
    const pts = directChild(lnSpc, NS.a, "spcPts");
    const pctVal = pct ? parseInt(pct.getAttribute("val") ?? "", 10) : NaN;
    const ptsVal = pts ? parseInt(pts.getAttribute("val") ?? "", 10) : NaN;
    if (!Number.isNaN(pctVal)) out.lnSpcPct = pctVal / 100000;
    else if (!Number.isNaN(ptsVal)) out.lnSpcPts = ptsVal / 100;
  }
  const spcPtsOf = (name: string): number | undefined => {
    const el = directChild(lvlPr, NS.a, name);
    const pts = el ? directChild(el, NS.a, "spcPts") : null;
    const v = pts ? parseInt(pts.getAttribute("val") ?? "", 10) : NaN;
    return Number.isNaN(v) ? undefined : v / 100;
  };
  const bef = spcPtsOf("spcBef");
  if (bef != null) out.spcBefPts = bef;
  const aft = spcPtsOf("spcAft");
  if (aft != null) out.spcAftPts = aft;

  const defRPr = directChild(lvlPr, NS.a, "defRPr");
  if (defRPr) {
    const sz = defRPr.getAttribute("sz");
    if (sz) {
      const n = parseInt(sz, 10);
      if (!Number.isNaN(n)) out.sizePt = n / 100;
    }
    const b = defRPr.getAttribute("b");
    if (b != null) out.bold = b === "1";
    const i = defRPr.getAttribute("i");
    if (i != null) out.italic = i === "1";
    const fill = directChild(defRPr, NS.a, "solidFill");
    if (fill) out.colorCss = resolveDrawingMlColor(parseColorChoice(fill), theme, clrMap);
    const latin = directChild(defRPr, NS.a, "latin");
    if (latin) out.fontFamily = resolveThemeFont(latin.getAttribute("typeface"), theme);
  }
  return out;
}

function mergeDefaults(base: RunStyleDefaults, lower: RunStyleDefaults): RunStyleDefaults {
  const out: RunStyleDefaults = { ...lower, ...base };
  // `...base` copies explicitly-undefined keys too; strip those so lower wins.
  for (const k of Object.keys(out) as (keyof RunStyleDefaults)[]) {
    if (out[k] === undefined && lower[k] !== undefined) {
      (out as Record<string, unknown>)[k] = lower[k];
    }
  }
  return out;
}

// Build a per-level resolver of default run styling for a shape, merging the
// full ECMA-376 inheritance chain: the shape's own txBody lstStyle, then the
// layout placeholder's lstStyle, the master placeholder's lstStyle, and
// finally the master's txStyles bucket for the placeholder type. Plain text
// boxes get only their own lstStyle.
export function buildDefaultsResolver(
  sp: Element,
  layoutDoc: Document | null,
  masterDoc: Document | null,
  theme: PptxTheme | null,
  clrMap: ClrMap,
): (level: number) => RunStyleDefaults {
  const sources: Element[] = [];
  const ownLst = lstStyleOf(sp);
  if (ownLst) sources.push(ownLst);

  const ph = getPh(sp);
  if (ph) {
    for (const doc of [layoutDoc, masterDoc]) {
      const match = findPhShape(doc, ph);
      const lst = match ? lstStyleOf(match) : null;
      if (lst) sources.push(lst);
    }
    const bucket = pickStyleEl(ph.type, masterDoc);
    if (bucket) sources.push(bucket);
  }

  const cache = new Map<number, RunStyleDefaults>();
  return (level: number): RunStyleDefaults => {
    const hit = cache.get(level);
    if (hit) return hit;
    let merged: RunStyleDefaults = {};
    for (const src of sources) {
      const lvlPr = nthLvlPr(src, level);
      if (lvlPr) merged = mergeDefaults(merged, readLvlPrDefaults(lvlPr, theme, clrMap));
    }
    cache.set(level, merged);
    return merged;
  };
}

function lstStyleOf(sp: Element): Element | null {
  const txBody = directChild(sp, NS.p, "txBody");
  return txBody ? directChild(txBody, NS.a, "lstStyle") : null;
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

// Resolve the effective a:bodyPr for a shape: its own txBody bodyPr, else the
// layout/master placeholder's. Individual attributes are read off the first
// element that defines them by the caller.
export function resolveBodyPrChain(
  sp: Element,
  layoutDoc: Document | null,
  masterDoc: Document | null,
): Element[] {
  const out: Element[] = [];
  const push = (shape: Element | null): void => {
    const txBody = shape ? directChild(shape, NS.p, "txBody") : null;
    const bodyPr = txBody ? directChild(txBody, NS.a, "bodyPr") : null;
    if (bodyPr) out.push(bodyPr);
  };
  push(sp);
  const ph = getPh(sp);
  if (ph) {
    push(findPhShape(layoutDoc, ph));
    push(findPhShape(masterDoc, ph));
  }
  return out;
}

// First defined attribute across a bodyPr chain.
export function bodyPrAttr(chain: Element[], name: string): string | null {
  for (const el of chain) {
    const v = el.getAttribute(name);
    if (v != null) return v;
  }
  return null;
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
