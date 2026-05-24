import { NS, firstChildNS, directChild } from "./ooxml";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import type { TextResolveCtx } from "./text";

// Resolve the slide background color, walking slide -> layout -> master and
// taking the first that defines one. Gradients are approximated by their first
// stop; image backgrounds are deferred (return null so the stage stays white).
export function resolveSlideBackground(
  slideDoc: Document,
  layoutDoc: Document | null,
  masterDoc: Document | null,
  ctx: TextResolveCtx,
): string | null {
  for (const doc of [slideDoc, layoutDoc, masterDoc]) {
    if (!doc) continue;
    const bg = findBg(doc);
    if (!bg) continue;
    const color = bgColor(bg, ctx);
    if (color) return color;
  }
  return null;
}

function findBg(doc: Document): Element | null {
  const cSld = doc.getElementsByTagNameNS(NS.p, "cSld")[0];
  if (!cSld) return null;
  return directChild(cSld, NS.p, "bg");
}

function bgColor(bg: Element, ctx: TextResolveCtx): string | null {
  const bgPr = directChild(bg, NS.p, "bgPr");
  if (bgPr) {
    const solid = directChild(bgPr, NS.a, "solidFill");
    if (solid) {
      const css = resolveDrawingMlColor(parseColorChoice(solid), ctx.theme, ctx.clrMap);
      if (css) return css;
    }
    const grad = directChild(bgPr, NS.a, "gradFill");
    if (grad) {
      const gs = firstChildNS(grad, NS.a, "gs"); // first gradient stop
      if (gs) {
        const css = resolveDrawingMlColor(parseColorChoice(gs), ctx.theme, ctx.clrMap);
        if (css) return css;
      }
    }
  }
  const bgRef = directChild(bg, NS.p, "bgRef");
  if (bgRef) {
    const css = resolveDrawingMlColor(parseColorChoice(bgRef), ctx.theme, ctx.clrMap);
    if (css) return css;
  }
  return null;
}
