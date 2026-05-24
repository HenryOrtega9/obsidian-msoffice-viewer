import { NS, directChild, firstChildNS, intAttr } from "./ooxml";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import type { TextResolveCtx } from "./text";

// Resolve a shape's fill to a CSS color, or null for noFill / unsupported
// fills. Gradients are approximated by their first stop (Phase 1 limitation
// carried forward). blipFill is a picture fill and is handled by images.ts.
export function resolveShapeFill(spPr: Element, ctx: TextResolveCtx): string | null {
  if (directChild(spPr, NS.a, "noFill")) return null;
  const solid = directChild(spPr, NS.a, "solidFill");
  if (solid) return resolveDrawingMlColor(parseColorChoice(solid), ctx.theme, ctx.clrMap);
  const grad = directChild(spPr, NS.a, "gradFill");
  if (grad) {
    const gs = firstChildNS(grad, NS.a, "gs");
    if (gs) return resolveDrawingMlColor(parseColorChoice(gs), ctx.theme, ctx.clrMap);
  }
  return null;
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
