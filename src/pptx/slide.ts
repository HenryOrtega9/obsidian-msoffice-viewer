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
import { getPh, resolvePlaceholderXfrm, buildPhDefaults } from "./inheritance";
import type { Chart } from "chart.js";
import { resolveShapeFill, resolveShapeLine } from "./shapes";
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
}

// Render one slide into `parent` as an absolutely-positioned stage. Walks the
// shape tree dispatching p:sp (shapes/text), p:pic (images), p:grpSp (nested
// groups) and p:cxnSp (connectors). p:graphicFrame (tables/charts) is added in
// later phases.
export async function renderSlide(
  slideRef: SlideRef,
  theme: PptxTheme | null,
  clrMap: ClrMap,
  scale: SlideScale,
  parent: HTMLElement,
  zip: JSZip,
  objectUrls: string[],
  charts: Chart[],
): Promise<void> {
  const ctx: RenderCtx = {
    theme,
    clrMap,
    zip,
    rels: slideRef.rels,
    objectUrls,
    charts,
    slideW: scale.widthPx,
    slideH: scale.heightPx,
  };

  const slideWrap = parent.createDiv({ cls: "docx-claude-pdf-slide" });
  const stage = slideWrap.createDiv({ cls: "docx-claude-pptx-slide" });
  stage.style.width = `${scale.widthPx}px`;
  stage.style.height = `${scale.heightPx}px`;

  const bg = resolveSlideBackground(slideRef.slideDoc, slideRef.layoutDoc, slideRef.masterDoc, ctx);
  if (bg) stage.style.background = bg;

  const spTree = findSpTree(slideRef.slideDoc);
  if (spTree) await walkSpTree(spTree, stage, ctx, slideRef, rootFrame(scale.scale));
}

async function walkSpTree(
  container: Element,
  stage: HTMLElement,
  ctx: RenderCtx,
  slideRef: SlideRef,
  frame: Frame,
): Promise<void> {
  for (const node of elementChildren(container)) {
    if (node.namespaceURI !== NS.p) continue;
    switch (node.localName) {
      case "sp":
        renderSp(node, stage, ctx, slideRef, frame);
        break;
      case "pic":
        await renderPic(node, stage, ctx, frame);
        break;
      case "grpSp":
        await renderGrp(node, stage, ctx, slideRef, frame);
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
  const box = shapeBox(sp, slideRef, frame, ctx);
  if (!box) return;

  const el = stage.createDiv({ cls: "docx-claude-pptx-shape" });
  setBox(el, box);

  const spPr = directChild(sp, NS.p, "spPr");
  if (spPr) {
    applyRotation(el, spPr);
    const fill = resolveShapeFill(spPr, ctx);
    if (fill) el.style.background = fill;
    const line = resolveShapeLine(spPr, ctx);
    if (line) {
      const w = line.widthEmu > 0 ? Math.max(0.5, line.widthEmu * frame.scaleX) : 1;
      el.style.border = `${w}px solid ${line.colorCss ?? "#000000"}`;
    }
  }

  const txBody = directChild(sp, NS.p, "txBody");
  if (txBody) {
    // Placeholder shapes inherit run defaults (size/color/font) from the
    // master's txStyles; plain text boxes use literal sizing.
    const ph = getPh(sp);
    const defaultsFor = ph
      ? buildPhDefaults(ph.type, slideRef.masterDoc, ctx.theme, ctx.clrMap)
      : undefined;
    renderParagraphsInto(parseTxBody(txBody, ctx, defaultsFor), el, frame.scaleX);
  }
}

async function renderPic(pic: Element, stage: HTMLElement, ctx: RenderCtx, frame: Frame): Promise<void> {
  const spPr = directChild(pic, NS.p, "spPr");
  const box = spPr ? boxFromXfrm(directChild(spPr, NS.a, "xfrm"), frame) : null;
  if (!box) return;
  await renderPicInto(pic, box, ctx.rels, ctx.zip, stage, ctx.objectUrls);
}

async function renderGrp(
  grpSp: Element,
  stage: HTMLElement,
  ctx: RenderCtx,
  slideRef: SlideRef,
  frame: Frame,
): Promise<void> {
  const grpSpPr = directChild(grpSp, NS.p, "grpSpPr");
  const xfrm = grpSpPr ? directChild(grpSpPr, NS.a, "xfrm") : null;
  const off = xfrm ? directChild(xfrm, NS.a, "off") : null;
  const ext = xfrm ? directChild(xfrm, NS.a, "ext") : null;
  const chOff = xfrm ? directChild(xfrm, NS.a, "chOff") : null;
  const chExt = xfrm ? directChild(xfrm, NS.a, "chExt") : null;

  // With a full group transform, build a child frame; otherwise recurse with
  // the parent frame so nested shapes still render at their absolute coords.
  let childFr = frame;
  if (off && ext && chOff && chExt) {
    const groupBox = boxInFrame(
      { x: intAttr(off, "x"), y: intAttr(off, "y") },
      { cx: intAttr(ext, "cx"), cy: intAttr(ext, "cy") },
      frame,
    );
    childFr = childFrame(
      groupBox,
      { x: intAttr(chOff, "x"), y: intAttr(chOff, "y") },
      { cx: intAttr(chExt, "cx"), cy: intAttr(chExt, "cy") },
    );
  }
  await walkSpTree(grpSp, stage, ctx, slideRef, childFr);
}

// Straight connectors rendered as an SVG line across the bounding box, honoring
// flipH/flipV to pick the diagonal direction.
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
  if (xfrm) applyRotation(el, xfrm);

  const flipV = xfrm?.getAttribute("flipV") === "1";
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "100%");
  svg.setAttribute("height", "100%");
  svg.setAttribute("preserveAspectRatio", "none");
  const lineEl = document.createElementNS("http://www.w3.org/2000/svg", "line");
  lineEl.setAttribute("x1", "0");
  lineEl.setAttribute("y1", flipV ? "100%" : "0");
  lineEl.setAttribute("x2", "100%");
  lineEl.setAttribute("y2", flipV ? "0" : "100%");
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

function applyRotation(el: HTMLElement, xfrm: Element): void {
  const rot = intAttr(xfrm, "rot", 0);
  if (rot) {
    el.style.transform = `rotate(${degFromRot(rot)}deg)`;
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

// Resolve a shape's pixel box: local xfrm, else inherited placeholder geometry
// (layout then master), else a default region by placeholder type so text
// still appears.
function shapeBox(sp: Element, slideRef: SlideRef, frame: Frame, ctx: RenderCtx): Box | null {
  const spPr = directChild(sp, NS.p, "spPr");
  let xfrm = spPr ? directChild(spPr, NS.a, "xfrm") : null;
  if (!xfrm) xfrm = resolvePlaceholderXfrm(sp, slideRef.layoutDoc, slideRef.masterDoc);
  const box = boxFromXfrm(xfrm, frame);
  if (box) return box;

  const ph = getPh(sp);
  if (ph) return defaultPhBox(ph.type, ctx.slideW, ctx.slideH);
  return null;
}

function defaultPhBox(type: string, slideW: number, slideH: number): Box {
  const isTitle = type === "title" || type === "ctrTitle";
  const frac = isTitle
    ? { l: 0.06, t: 0.05, w: 0.88, h: 0.2 }
    : { l: 0.06, t: 0.3, w: 0.88, h: 0.6 };
  return { left: frac.l * slideW, top: frac.t * slideH, width: frac.w * slideW, height: frac.h * slideH };
}
