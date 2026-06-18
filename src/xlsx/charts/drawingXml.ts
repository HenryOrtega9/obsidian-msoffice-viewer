import { type AnchorPoint, EMU_PER_PX } from "../geometry";
import { NS, firstChildNS, relId } from "./ooxml";

export interface DrawingChartAnchor {
  from: AnchorPoint;
  to: AnchorPoint | null;
  // Explicit pixel size for a oneCellAnchor (from + ext). null for twoCellAnchor
  // (where `to` drives the size). Stored in px to match anchorRangeToBox.
  ext: { width: number; height: number } | null;
  // absoluteAnchor box in px (pos + ext, EMU→px), positioned independent of
  // cells. null for cell-based anchors; when set it overrides from/to/ext.
  abs: { x: number; y: number; width: number; height: number } | null;
  chartRelId: string;
}

// Find every graphicFrame in a drawing that references a chart, returning its
// anchor box and the relationship id pointing at the chart part.
export function parseDrawingForCharts(doc: Document): DrawingChartAnchor[] {
  const out: DrawingChartAnchor[] = [];
  const root = doc.documentElement;
  if (!root) return out;

  for (const anchor of Array.from(root.childNodes)) {
    if (anchor.nodeType !== 1) continue;
    const el = anchor as Element;
    if (el.namespaceURI !== NS.xdr) continue;
    if (
      el.localName !== "twoCellAnchor" &&
      el.localName !== "oneCellAnchor" &&
      el.localName !== "absoluteAnchor"
    ) {
      continue;
    }

    const frame = firstChildNS(el, NS.xdr, "graphicFrame");
    if (!frame) continue;
    const graphicData = frame.getElementsByTagNameNS(NS.a, "graphicData")[0];
    if (!graphicData) continue;
    const chartEl = graphicData.getElementsByTagNameNS(NS.c, "chart")[0];
    if (!chartEl) continue;
    const chartRelId = relId(chartEl);
    if (!chartRelId) continue;

    if (el.localName === "absoluteAnchor") {
      // pos + ext (both EMU), positioned independent of the grid.
      const abs = parseAbsBox(el);
      if (!abs) continue;
      out.push({ from: { col: 0, colOff: 0, row: 0, rowOff: 0 }, to: null, ext: null, abs, chartRelId });
      continue;
    }

    const from = parseAnchorPoint(firstChildNS(el, NS.xdr, "from"));
    if (!from) continue;
    const to = parseAnchorPoint(firstChildNS(el, NS.xdr, "to"));
    // oneCellAnchor sizes via a direct xdr:ext (cx/cy in EMU); twoCellAnchor has
    // none (and the inner xfrm a:ext is namespace-disambiguated, so it won't match).
    const ext = parseExt(el);
    out.push({ from, to, ext, abs: null, chartRelId });
  }
  return out;
}

function parseAbsBox(
  el: Element,
): { x: number; y: number; width: number; height: number } | null {
  const pos = firstChildNS(el, NS.xdr, "pos");
  const ext = firstChildNS(el, NS.xdr, "ext");
  if (!pos || !ext) return null;
  const x = parseInt(pos.getAttribute("x") ?? "", 10);
  const y = parseInt(pos.getAttribute("y") ?? "", 10);
  const cx = parseInt(ext.getAttribute("cx") ?? "", 10);
  const cy = parseInt(ext.getAttribute("cy") ?? "", 10);
  if (![x, y, cx, cy].every((n) => Number.isFinite(n)) || cx <= 0 || cy <= 0) return null;
  return { x: x / EMU_PER_PX, y: y / EMU_PER_PX, width: cx / EMU_PER_PX, height: cy / EMU_PER_PX };
}

function parseExt(el: Element): { width: number; height: number } | null {
  const extEl = firstChildNS(el, NS.xdr, "ext");
  if (!extEl) return null;
  const cx = parseInt(extEl.getAttribute("cx") ?? "", 10);
  const cy = parseInt(extEl.getAttribute("cy") ?? "", 10);
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) return null;
  return { width: cx / EMU_PER_PX, height: cy / EMU_PER_PX };
}

function parseAnchorPoint(el: Element | null): AnchorPoint | null {
  if (!el) return null;
  const col = intText(el, "col");
  const colOff = intText(el, "colOff");
  const row = intText(el, "row");
  const rowOff = intText(el, "rowOff");
  if (col == null || row == null) return null;
  return { col, colOff: colOff ?? 0, row, rowOff: rowOff ?? 0 };
}

function intText(parent: Element, local: string): number | null {
  const el = firstChildNS(parent, NS.xdr, local);
  if (!el || el.textContent == null) return null;
  const n = parseInt(el.textContent, 10);
  return Number.isFinite(n) ? n : null;
}
