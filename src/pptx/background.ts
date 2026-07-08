import { NS, directChild, intAttr } from "./ooxml";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import { resolveFillCss } from "./shapes";
import type { TextResolveCtx } from "./text";

// Resolve the slide background, walking slide -> layout -> master and stopping
// at the FIRST document that defines a p:bg — even if that background can't be
// fully resolved (e.g. a picture fill). Falling through past an unsupported
// slide background would paint the master's color behind text authored for a
// different backdrop.
export function resolveSlideBackground(
  slideDoc: Document,
  layoutDoc: Document | null,
  masterDoc: Document | null,
  ctx: TextResolveCtx,
): string | null {
  for (const doc of [slideDoc, layoutDoc, masterDoc]) {
    if (!doc) continue;
    const bg = findBg(doc);
    if (bg) return bgColor(bg, ctx);
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
  if (bgPr) return resolveFillCss(bgPr, ctx);

  // bgRef: idx >= 1000 indexes the theme's bgFillStyleLst (1001-based); the
  // ref's own color substitutes for phClr. idx 0/999-below falls back to the
  // ref color itself.
  const bgRef = directChild(bg, NS.p, "bgRef");
  if (bgRef) {
    const refCss = resolveDrawingMlColor(parseColorChoice(bgRef), ctx.theme, ctx.clrMap);
    const idx = intAttr(bgRef, "idx", 0);
    if (idx >= 1000 && refCss && /^#[0-9A-Fa-f]{6}$/.test(refCss)) {
      const entry = ctx.theme?.bgFillStyles?.[idx - 1001];
      if (entry) {
        const css = resolveFillCss(wrapEntry(entry), ctx, refCss.slice(1).toUpperCase());
        if (css) return css;
      }
    }
    return refCss;
  }
  return null;
}

// resolveFillCss scans a parent's children for the fill choice; a fmtScheme
// entry IS the fill element, so hand it a lightweight wrapper.
function wrapEntry(entry: Element): Element {
  const parent = entry.ownerDocument.createElementNS(NS.p, "bgWrap");
  parent.appendChild(entry.cloneNode(true));
  return parent;
}
