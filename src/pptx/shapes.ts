import { NS, directChild, childrenNS, intAttr } from "./ooxml";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import type { TextResolveCtx } from "./text";

// Resolve a fill-bearing element's children (a:solidFill / a:gradFill /
// a:pattFill / a:noFill) to a CSS background value. Gradients become real CSS
// linear-gradients with stops sorted by position; pattern fills approximate to
// their foreground color. `phClr` substitutes into scheme val="phClr" refs
// when resolving theme format-scheme style entries.
export function resolveFillCss(
  parent: Element,
  ctx: TextResolveCtx,
  phClr?: string | null,
): string | null {
  if (directChild(parent, NS.a, "noFill")) return null;
  const solid = directChild(parent, NS.a, "solidFill");
  if (solid) return resolveDrawingMlColor(parseColorChoice(solid), ctx.theme, ctx.clrMap, phClr);
  const grad = directChild(parent, NS.a, "gradFill");
  if (grad) return gradientCss(grad, ctx, phClr);
  const patt = directChild(parent, NS.a, "pattFill");
  if (patt) {
    const fg = directChild(patt, NS.a, "fgClr");
    if (fg) return resolveDrawingMlColor(parseColorChoice(fg), ctx.theme, ctx.clrMap, phClr);
  }
  return null;
}

// Build a CSS linear-gradient from a:gradFill. Stops are sorted by pos (the
// gsLst need not be authored in order). OOXML's lin angle is clockwise from
// 3 o'clock; CSS gradient angle is clockwise from 12 o'clock, so add 90deg.
function gradientCss(grad: Element, ctx: TextResolveCtx, phClr?: string | null): string | null {
  const gsLst = directChild(grad, NS.a, "gsLst");
  if (!gsLst) return null;
  const stops: { pos: number; css: string }[] = [];
  for (const gs of childrenNS(gsLst, NS.a, "gs")) {
    const css = resolveDrawingMlColor(parseColorChoice(gs), ctx.theme, ctx.clrMap, phClr);
    if (!css) continue;
    stops.push({ pos: intAttr(gs, "pos", 0) / 100000, css });
  }
  if (stops.length === 0) return null;
  stops.sort((a, b) => a.pos - b.pos);
  if (stops.length === 1) return stops[0].css;

  const lin = directChild(grad, NS.a, "lin");
  const angDeg = lin ? intAttr(lin, "ang", 0) / 60000 : 0;
  const cssAngle = (angDeg + 90) % 360;
  const stopList = stops.map((s) => `${s.css} ${(s.pos * 100).toFixed(1)}%`).join(", ");
  return `linear-gradient(${cssAngle}deg, ${stopList})`;
}

// Resolve a shape's explicit spPr fill. Returns null for noFill or when spPr
// carries no fill choice (callers then consult the p:style fillRef).
export function resolveShapeFill(spPr: Element, ctx: TextResolveCtx): string | null {
  return resolveFillCss(spPr, ctx);
}

// Distinguish "spPr says noFill" from "spPr is silent" so the style ref is
// only consulted in the latter case.
export function hasExplicitFillChoice(spPr: Element): boolean {
  return !!(
    directChild(spPr, NS.a, "noFill") ||
    directChild(spPr, NS.a, "solidFill") ||
    directChild(spPr, NS.a, "gradFill") ||
    directChild(spPr, NS.a, "pattFill") ||
    directChild(spPr, NS.a, "blipFill") ||
    directChild(spPr, NS.a, "grpFill")
  );
}

export interface ShapeLine {
  colorCss: string | null;
  widthEmu: number; // a:ln/@w, 0 if unspecified
}

export function resolveShapeLine(spPr: Element, ctx: TextResolveCtx): ShapeLine | null {
  const ln = directChild(spPr, NS.a, "ln");
  if (!ln) return null;
  if (directChild(ln, NS.a, "noFill")) return null;
  const solid = directChild(ln, NS.a, "solidFill");
  const colorCss = solid ? resolveDrawingMlColor(parseColorChoice(solid), ctx.theme, ctx.clrMap) : null;
  const widthEmu = intAttr(ln, "w", 0);
  if (!colorCss && widthEmu === 0) return null;
  return { colorCss, widthEmu };
}

export function shapeLineIsNoFill(spPr: Element): boolean {
  const ln = directChild(spPr, NS.a, "ln");
  return !!ln && !!directChild(ln, NS.a, "noFill");
}

// Styling resolved from a shape's p:style element (theme references used by
// default PowerPoint-inserted shapes, which carry no explicit spPr fill/line).
export interface StyleRefResult {
  fillCss: string | null;
  line: ShapeLine | null;
  fontColorCss: string | null;
}

// Resolve p:style: fillRef indexes the theme fmtScheme fillStyleLst (1-based;
// >=1000 indexes bgFillStyleLst), lnRef the lnStyleLst, and each ref's own
// color substitutes for phClr inside the referenced style. fontRef supplies
// the default run color.
export function resolveStyleRef(sp: Element, ctx: TextResolveCtx): StyleRefResult | null {
  const style = directChild(sp, NS.p, "style");
  if (!style) return null;
  const out: StyleRefResult = { fillCss: null, line: null, fontColorCss: null };

  const refColor = (ref: Element): string | null => {
    const c = resolveDrawingMlColor(parseColorChoice(ref), ctx.theme, ctx.clrMap);
    return c;
  };
  // phClr substitution needs a plain "RRGGBB"; strip a leading '#'. rgba()
  // strings (alpha-modified refs) fall back to the raw ref color as the fill.
  const asHex = (css: string | null): string | null =>
    css && /^#[0-9A-Fa-f]{6}$/.test(css) ? css.slice(1).toUpperCase() : null;

  const fillRef = directChild(style, NS.p, "fillRef");
  if (fillRef) {
    const idx = intAttr(fillRef, "idx", 0);
    const css = refColor(fillRef);
    const hex = asHex(css);
    const styles = idx >= 1000 ? ctx.theme?.bgFillStyles : ctx.theme?.fillStyles;
    const styleEl = styles?.[(idx >= 1000 ? idx - 1000 : idx) - 1];
    if (idx > 0 && styleEl && hex) {
      out.fillCss = resolveFillFromStyleEntry(styleEl, ctx, hex) ?? css;
    } else if (idx > 0) {
      out.fillCss = css;
    }
  }

  const lnRef = directChild(style, NS.p, "lnRef");
  if (lnRef) {
    const idx = intAttr(lnRef, "idx", 0);
    const css = refColor(lnRef);
    if (idx > 0 && css) {
      const lnEl = ctx.theme?.lnStyles?.[idx - 1];
      const widthEmu = lnEl ? intAttr(lnEl, "w", 9525) : 9525;
      out.line = { colorCss: css, widthEmu };
    }
  }

  const fontRef = directChild(style, NS.p, "fontRef");
  if (fontRef) out.fontColorCss = refColor(fontRef);

  return out;
}

// A fmtScheme style entry is itself a fill element (a:solidFill / a:gradFill /
// a:pattFill) whose colors reference phClr.
function resolveFillFromStyleEntry(
  styleEl: Element,
  ctx: TextResolveCtx,
  phClrHex: string,
): string | null {
  switch (styleEl.localName) {
    case "solidFill":
      return resolveDrawingMlColor(parseColorChoice(styleEl), ctx.theme, ctx.clrMap, phClrHex);
    case "gradFill":
      return gradientCss(styleEl, ctx, phClrHex);
    case "pattFill": {
      const fg = directChild(styleEl, NS.a, "fgClr");
      return fg ? resolveDrawingMlColor(parseColorChoice(fg), ctx.theme, ctx.clrMap, phClrHex) : null;
    }
    default:
      return null;
  }
}
