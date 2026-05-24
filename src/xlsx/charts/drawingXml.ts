import type { AnchorPoint } from "../geometry";
import { NS, firstChildNS, relId } from "./ooxml";

export interface DrawingChartAnchor {
  from: AnchorPoint;
  to: AnchorPoint | null;
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
    if (el.localName !== "twoCellAnchor" && el.localName !== "oneCellAnchor") continue;

    const frame = firstChildNS(el, NS.xdr, "graphicFrame");
    if (!frame) continue;
    const graphicData = frame.getElementsByTagNameNS(NS.a, "graphicData")[0];
    if (!graphicData) continue;
    const chartEl = graphicData.getElementsByTagNameNS(NS.c, "chart")[0];
    if (!chartEl) continue;
    const chartRelId = relId(chartEl);
    if (!chartRelId) continue;

    const from = parseAnchorPoint(firstChildNS(el, NS.xdr, "from"));
    const to = parseAnchorPoint(firstChildNS(el, NS.xdr, "to"));
    if (!from) continue;
    out.push({ from, to, chartRelId });
  }
  return out;
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
