import type JSZip from "jszip";
import { NS, directChild, elementChildren, intAttr } from "./ooxml";
import {
  type Box,
  type SlideScale,
  type Frame,
  rootFrame,
  boxInFrame,
  childFrame,
  degFromRot,
} from "./geometry";
import { parseTxBody, renderParagraphsInto, type TextResolveCtx } from "./text";
import {
  getPh,
  resolvePlaceholderXfrm,
  buildDefaultsResolver,
  resolveBodyPrChain,
  bodyPrAttr,
} from "./inheritance";
import type { Chart } from "chart.js";
import {
  resolveShapeFill,
  resolveShapeLine,
  resolveStyleRef,
  hasExplicitFillChoice,
  shapeLineIsNoFill,
} from "./shapes";
import { renderPicInto } from "./images";
import { renderTableInto } from "./tables";
import { renderChartFrame } from "./charts/render";
import { resolveSlideBackground } from "./background";
import { findSpTree, type SlideRef } from "./presentation";
import type { PptxTheme, ClrMap } from "./themes";

interface RenderCtx extends TextResolveCtx {
  zip: JSZip;
  rels: Map<string, string>;
  objectUrls: string[];
  charts: Chart[];
  slideW: number;
  slideH: number;
  tableStyles: Document | null;
}

// Render one slide into `parent` as an absolutely-positioned stage. Per
// ECMA-376, the master's and layout's non-placeholder shapes (logos, footer
// bars, decorations) render beneath the slide's own tree unless the slide (or
// layout) sets showMasterSp="0"; placeholder shapes on master/layout are
// prototypes only and never render themselves.
export async function renderSlide(
  slideRef: SlideRef,
  theme: PptxTheme | null,
  clrMap: ClrMap,
  scale: SlideScale,
  parent: HTMLElement,
  zip: JSZip,
  objectUrls: string[],
  charts: Chart[],
  tableStyles: Document | null,
): Promise<void> {
  const baseCtx = {
    theme,
    clrMap,
    zip,
    objectUrls,
    charts,
    slideW: scale.widthPx,
    slideH: scale.heightPx,
    tableStyles,
  };

  const slideWrap = parent.createDiv({ cls: "docx-claude-pdf-slide" });
  const stage = slideWrap.createDiv({ cls: "docx-claude-pptx-slide" });
  stage.style.width = `${scale.widthPx}px`;
  stage.style.height = `${scale.heightPx}px`;

  const slideCtx: RenderCtx = { ...baseCtx, rels: slideRef.rels };
  const bg = resolveSlideBackground(slideRef.slideDoc, slideRef.layoutDoc, slideRef.masterDoc, slideCtx);
  if (bg) stage.style.background = bg;

  if (showMasterShapes(slideRef)) {
    const inherited: [Document | null, Map<string, string>][] = [
      [slideRef.masterDoc, slideRef.masterRels],
      [slideRef.layoutDoc, slideRef.layoutRels],
    ];
    for (const [doc, rels] of inherited) {
      if (!doc) continue;
      const tree = findSpTree(doc);
      if (!tree) continue;
      await walkSpTree(tree, stage, { ...baseCtx, rels }, slideRef, rootFrame(scale.scale), true);
    }
  }

  const spTree = findSpTree(slideRef.slideDoc);
  if (spTree) await walkSpTree(spTree, stage, slideCtx, slideRef, rootFrame(scale.scale), false);
}

// showMasterSp defaults to on; the slide's attribute wins over the layout's.
function showMasterShapes(slideRef: SlideRef): boolean {
  const slideAttr = slideRef.slideDoc.documentElement.getAttribute("showMasterSp");
  if (slideAttr != null) return slideAttr !== "0";
  const layoutAttr = slideRef.layoutDoc?.documentElement.getAttribute("showMasterSp");
  if (layoutAttr != null) return layoutAttr !== "0";
  return true;
}

async function walkSpTree(
  container: Element,
  stage: HTMLElement,
  ctx: RenderCtx,
  slideRef: SlideRef,
  frame: Frame,
  skipPlaceholders: boolean,
): Promise<void> {
  for (const node of elementChildren(container)) {
    if (node.namespaceURI !== NS.p) continue;
    switch (node.localName) {
      case "sp":
        if (skipPlaceholders && getPh(node)) break;
        renderSp(node, stage, ctx, slideRef, frame);
        break;
      case "pic":
        await renderPic(node, stage, ctx, frame);
        break;
      case "grpSp":
        await renderGrp(node, stage, ctx, slideRef, frame, skipPlaceholders);
        break;
      case "cxnSp":
        renderCxn(node, stage, ctx, frame);
        break;
      case "graphicFrame":
        await renderGraphicFrame(node, stage, ctx, frame);
        break;
    }
  }
}

function renderSp(
  sp: Element,
  stage: HTMLElement,
  ctx: RenderCtx,
  slideRef: SlideRef,
  frame: Frame,
): void {
  const spPr = directChild(sp, NS.p, "spPr");
  const localXfrm = spPr ? directChild(spPr, NS.a, "xfrm") : null;
  const effXfrm = localXfrm ?? resolvePlaceholderXfrm(sp, slideRef.layoutDoc, slideRef.masterDoc);
  let box = boxFromXfrm(effXfrm, frame);
  if (!box) {
    const ph = getPh(sp);
    if (!ph) return;
    box = defaultPhBox(ph.type, ctx.slideW, ctx.slideH);
  }

  const el = stage.createDiv({ cls: "docx-claude-pptx-shape" });
  setBox(el, box);
  applyTransform(el, effXfrm);

  // Explicit spPr fill/line first; a silent spPr falls back to the shape's
  // p:style theme references (how default-inserted shapes are colored).
  const styleRef = resolveStyleRef(sp, ctx);
  if (spPr) {
    let fill = resolveShapeFill(spPr, ctx);
    if (fill == null && !hasExplicitFillChoice(spPr) && styleRef?.fillCss) {
      fill = styleRef.fillCss;
    }
    if (fill) el.style.background = fill;

    let line = resolveShapeLine(spPr, ctx);
    if (!line && !shapeLineIsNoFill(spPr) && styleRef?.line) line = styleRef.line;
    if (line) {
      const w = line.widthEmu > 0 ? Math.max(0.5, line.widthEmu * frame.scaleX) : 1;
      el.style.border = `${w}px solid ${line.colorCss ?? "#000000"}`;
    }
    applyPresetGeometry(el, spPr, box);
  }

  const txBody = directChild(sp, NS.p, "txBody");
  if (txBody) {
    // Full list-style inheritance chain (shape -> layout ph -> master ph ->
    // master txStyles), with the style fontRef color as the last fallback.
    const resolver = buildDefaultsResolver(sp, slideRef.layoutDoc, slideRef.masterDoc, ctx.theme, ctx.clrMap);
    const defaultsFor = (level: number) => {
      const d = resolver(level);
      if (d.colorCss == null && styleRef?.fontColorCss) {
        return { ...d, colorCss: styleRef.fontColorCss };
      }
      return d;
    };
    const bodyPrChain = resolveBodyPrChain(sp, slideRef.layoutDoc, slideRef.masterDoc);
    applyBodyInsets(el, bodyPrChain, frame.scaleX);
    applyBodyAnchor(el, bodyPrChain);
    const { fontScale, lnSpcReduction } = readAutofit(bodyPrChain);
    renderParagraphsInto(parseTxBody(txBody, ctx, defaultsFor), el, frame.scaleX, fontScale, lnSpcReduction);
  }
}

// Approximate preset geometry with CSS. Ellipses and rounded rectangles cover
// the bulk of real decks; a few straight-edged presets map to clip-path.
// Unknown presets keep the rectangular bounding box.
const CLIP_PRESETS: Record<string, string> = {
  triangle: "polygon(50% 0%, 100% 100%, 0% 100%)",
  rtTriangle: "polygon(0% 0%, 100% 100%, 0% 100%)",
  diamond: "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)",
  pentagon: "polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)",
  hexagon: "polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%)",
  parallelogram: "polygon(20% 0%, 100% 0%, 80% 100%, 0% 100%)",
  trapezoid: "polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%)",
  chevron: "polygon(0% 0%, 75% 0%, 100% 50%, 75% 100%, 0% 100%, 25% 50%)",
  homePlate: "polygon(0% 0%, 75% 0%, 100% 50%, 75% 100%, 0% 100%)",
};

function applyPresetGeometry(el: HTMLElement, spPr: Element, box: Box): void {
  const prstGeom = directChild(spPr, NS.a, "prstGeom");
  const prst = prstGeom?.getAttribute("prst");
  if (!prst) return;
  if (prst === "ellipse") {
    el.style.borderRadius = "50%";
    return;
  }
  if (prst === "roundRect" || prst === "round2SameRect" || prst === "snip1Rect") {
    // roundRect radius = adj (default 16.67%) of the shorter side.
    let adj = 16667 / 100000;
    const avLst = prstGeom ? directChild(prstGeom, NS.a, "avLst") : null;
    const gd = avLst ? directChild(avLst, NS.a, "gd") : null;
    const fmla = gd?.getAttribute("fmla") ?? "";
    const m = fmla.match(/^val (\d+)$/);
    if (m) adj = parseInt(m[1], 10) / 100000;
    const r = Math.min(box.width, box.height) * Math.max(0, Math.min(0.5, adj));
    el.style.borderRadius = `${r}px`;
    return;
  }
  const clip = CLIP_PRESETS[prst];
  if (clip) el.style.clipPath = clip;
}

// Vertical anchoring (bodyPr anchor: t/ctr/b) via flex column layout;
// paragraphs become flex items so justify-content positions the block.
function applyBodyAnchor(el: HTMLElement, chain: Element[]): void {
  const anchor = bodyPrAttr(chain, "anchor");
  if (anchor !== "ctr" && anchor !== "b") return;
  el.style.display = "flex";
  el.style.flexDirection = "column";
  el.style.justifyContent = anchor === "ctr" ? "center" : "flex-end";
}

// normAutofit carries the shrink PowerPoint baked in to fit text: fontScale
// scales run sizes, lnSpcReduction shrinks line spacing. Read from the first
// bodyPr in the chain that has one.
function readAutofit(chain: Element[]): { fontScale: number; lnSpcReduction: number } {
  for (const bodyPr of chain) {
    const af = directChild(bodyPr, NS.a, "normAutofit");
    if (!af) continue;
    const pct = (attr: string): number => {
      const v = af.getAttribute(attr);
      const n = v ? parseInt(v, 10) : NaN;
      return Number.isFinite(n) && n > 0 ? n / 100000 : attr === "fontScale" ? 1 : 0;
    };
    return { fontScale: pct("fontScale"), lnSpcReduction: pct("lnSpcReduction") };
  }
  return { fontScale: 1, lnSpcReduction: 0 };
}

// Apply text-frame insets as padding so text sits inside the box like PPT.
// Defaults are the OOXML defaults (0.1" left/right, 0.05" top/bottom); each
// side reads the first bodyPr in the inheritance chain that specifies it.
function applyBodyInsets(el: HTMLElement, chain: Element[], scale: number): void {
  const ins = (attr: string, def: number): number => {
    const v = bodyPrAttr(chain, attr);
    const emu = v != null ? parseInt(v, 10) : def;
    return Number.isFinite(emu) ? Math.max(0, emu) * scale : 0;
  };
  el.style.boxSizing = "border-box";
  el.style.paddingLeft = `${ins("lIns", 91440)}px`;
  el.style.paddingRight = `${ins("rIns", 91440)}px`;
  el.style.paddingTop = `${ins("tIns", 45720)}px`;
  el.style.paddingBottom = `${ins("bIns", 45720)}px`;
}

async function renderPic(pic: Element, stage: HTMLElement, ctx: RenderCtx, frame: Frame): Promise<void> {
  const spPr = directChild(pic, NS.p, "spPr");
  const xfrm = spPr ? directChild(spPr, NS.a, "xfrm") : null;
  const box = boxFromXfrm(xfrm, frame);
  if (!box) return;
  const holder = await renderPicInto(pic, box, ctx.rels, ctx.zip, stage, ctx.objectUrls);
  if (holder) applyTransform(holder, xfrm);
}

async function renderGrp(
  grpSp: Element,
  stage: HTMLElement,
  ctx: RenderCtx,
  slideRef: SlideRef,
  frame: Frame,
  skipPlaceholders: boolean,
): Promise<void> {
  const grpSpPr = directChild(grpSp, NS.p, "grpSpPr");
  const xfrm = grpSpPr ? directChild(grpSpPr, NS.a, "xfrm") : null;
  const off = xfrm ? directChild(xfrm, NS.a, "off") : null;
  const ext = xfrm ? directChild(xfrm, NS.a, "ext") : null;
  const chOff = xfrm ? directChild(xfrm, NS.a, "chOff") : null;
  const chExt = xfrm ? directChild(xfrm, NS.a, "chExt") : null;

  // With a full group transform, render children into a wrapper positioned at
  // the group box; the wrapper carries the group's rotate/flip so children
  // transform together. Otherwise recurse with the parent frame so nested
  // shapes still render at their absolute coords.
  if (off && ext && chOff && chExt) {
    const groupBox = boxInFrame(
      { x: intAttr(off, "x"), y: intAttr(off, "y") },
      { cx: intAttr(ext, "cx"), cy: intAttr(ext, "cy") },
      frame,
    );
    const wrapper = stage.createDiv({ cls: "docx-claude-pptx-group" });
    setBox(wrapper, groupBox);
    applyTransform(wrapper, xfrm);
    const childFr = childFrame(
      { left: 0, top: 0, width: groupBox.width, height: groupBox.height },
      { x: intAttr(chOff, "x"), y: intAttr(chOff, "y") },
      { cx: intAttr(chExt, "cx"), cy: intAttr(chExt, "cy") },
    );
    await walkSpTree(grpSp, wrapper, ctx, slideRef, childFr, skipPlaceholders);
    return;
  }
  await walkSpTree(grpSp, stage, ctx, slideRef, frame, skipPlaceholders);
}

// Straight connectors rendered as an SVG line across the bounding box. flipH
// and flipV each independently toggle which diagonal the line follows.
function renderCxn(cxnSp: Element, stage: HTMLElement, ctx: RenderCtx, frame: Frame): void {
  const spPr = directChild(cxnSp, NS.p, "spPr");
  const xfrm = spPr ? directChild(spPr, NS.a, "xfrm") : null;
  const box = boxFromXfrm(xfrm, frame);
  if (!box || !spPr) return;
  const line = resolveShapeLine(spPr, ctx);
  const color = line?.colorCss ?? "#000000";
  const w = line && line.widthEmu > 0 ? Math.max(0.75, line.widthEmu * frame.scaleX) : 1;

  const el = stage.createDiv({ cls: "docx-claude-pptx-shape" });
  setBox(el, box);
  if (xfrm) {
    const rot = intAttr(xfrm, "rot", 0);
    if (rot) {
      el.style.transform = `rotate(${degFromRot(rot)}deg)`;
      el.style.transformOrigin = "center center";
    }
  }

  const flipH = xfrm?.getAttribute("flipH") === "1";
  const flipV = xfrm?.getAttribute("flipV") === "1";
  const downRight = flipH === flipV; // both flips (or neither) keep the ↘ diagonal
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "100%");
  svg.setAttribute("height", "100%");
  svg.setAttribute("preserveAspectRatio", "none");
  const lineEl = document.createElementNS("http://www.w3.org/2000/svg", "line");
  lineEl.setAttribute("x1", "0");
  lineEl.setAttribute("y1", downRight ? "0" : "100%");
  lineEl.setAttribute("x2", "100%");
  lineEl.setAttribute("y2", downRight ? "100%" : "0");
  lineEl.setAttribute("stroke", color);
  lineEl.setAttribute("stroke-width", String(w));
  svg.appendChild(lineEl);
  el.appendChild(svg);
}

// p:graphicFrame holds tables and charts. Note it carries a p:xfrm (not the
// a:xfrm shapes use), though its a:off/a:ext children are still drawingml.
async function renderGraphicFrame(
  gf: Element,
  stage: HTMLElement,
  ctx: RenderCtx,
  frame: Frame,
): Promise<void> {
  const box = boxFromXfrm(directChild(gf, NS.p, "xfrm"), frame);
  if (!box) return;
  const graphic = directChild(gf, NS.a, "graphic");
  const graphicData = graphic ? directChild(graphic, NS.a, "graphicData") : null;
  if (!graphicData) return;
  const uri = graphicData.getAttribute("uri") ?? "";
  if (uri.includes("/table")) {
    renderTableInto(gf, box, ctx, frame.scaleX, stage);
  } else if (uri.includes("/chart")) {
    const chart = await renderChartFrame(graphicData, box, ctx.zip, ctx.rels, ctx.theme, stage);
    if (chart) ctx.charts.push(chart);
  }
}

// Compose rotation and flips from an a:xfrm into one CSS transform about the
// shape's center.
function applyTransform(el: HTMLElement, xfrm: Element | null): void {
  if (!xfrm) return;
  const parts: string[] = [];
  const rot = intAttr(xfrm, "rot", 0);
  if (rot) parts.push(`rotate(${degFromRot(rot)}deg)`);
  if (xfrm.getAttribute("flipH") === "1") parts.push("scaleX(-1)");
  if (xfrm.getAttribute("flipV") === "1") parts.push("scaleY(-1)");
  if (parts.length > 0) {
    el.style.transform = parts.join(" ");
    el.style.transformOrigin = "center center";
  }
}

function boxFromXfrm(xfrm: Element | null, frame: Frame): Box | null {
  if (!xfrm) return null;
  const off = directChild(xfrm, NS.a, "off");
  const ext = directChild(xfrm, NS.a, "ext");
  if (!off || !ext) return null;
  return boxInFrame(
    { x: intAttr(off, "x"), y: intAttr(off, "y") },
    { cx: intAttr(ext, "cx"), cy: intAttr(ext, "cy") },
    frame,
  );
}

function setBox(el: HTMLElement, box: Box): void {
  el.style.left = `${box.left}px`;
  el.style.top = `${box.top}px`;
  el.style.width = `${box.width}px`;
  el.style.height = `${box.height}px`;
}

function defaultPhBox(type: string, slideW: number, slideH: number): Box {
  const isTitle = type === "title" || type === "ctrTitle";
  const frac = isTitle
    ? { l: 0.06, t: 0.05, w: 0.88, h: 0.2 }
    : { l: 0.06, t: 0.3, w: 0.88, h: 0.6 };
  return { left: frac.l * slideW, top: frac.t * slideH, width: frac.w * slideW, height: frac.h * slideH };
}
