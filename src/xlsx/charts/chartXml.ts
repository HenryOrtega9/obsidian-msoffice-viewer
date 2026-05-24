import { ChartKind, ChartSeries, ChartSpec } from "./types";
import { NS, firstChildNS } from "./ooxml";

const TYPE_TO_KIND: Record<string, ChartKind> = {
  barChart: "bar",
  bar3DChart: "bar",
  lineChart: "line",
  line3DChart: "line",
  pieChart: "pie",
  pie3DChart: "pie",
  doughnutChart: "doughnut",
  scatterChart: "scatter",
  areaChart: "area",
  area3DChart: "area",
  radarChart: "radar",
  bubbleChart: "bubble",
};

const SCHEME_TO_THEME_INDEX: Record<string, number> = {
  dk1: 0, tx1: 0,
  lt1: 1, bg1: 1,
  dk2: 2, tx2: 2,
  lt2: 3, bg2: 3,
  accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9,
  hlink: 10, folHlink: 11,
};

export function parseChartXml(doc: Document, theme?: readonly string[]): ChartSpec | null {
  const root = doc.documentElement;
  if (!root) return null;
  const chart = firstChildNS(root, NS.c, "chart");
  if (!chart) return null;
  const plotArea = firstChildNS(chart, NS.c, "plotArea");
  if (!plotArea) return null;

  const typeEl = findChartTypeElement(plotArea);
  if (!typeEl) return null;
  const localName = typeEl.localName;
  const kind = TYPE_TO_KIND[localName] ?? "unsupported";

  const spec: ChartSpec = {
    kind,
    rawType: kind === "unsupported" ? localName : undefined,
    title: parseTitle(chart),
    series: [],
  };

  if (kind === "bar") {
    const barDir = attrVal(firstChildNS(typeEl, NS.c, "barDir"));
    spec.barHorizontal = barDir === "bar";
  }
  if (kind === "bar" || kind === "line" || kind === "area") {
    const grouping = attrVal(firstChildNS(typeEl, NS.c, "grouping"));
    spec.stacked = grouping === "stacked" || grouping === "percentStacked";
  }

  if (kind === "unsupported") return spec;

  const serEls = childrenNS(typeEl, NS.c, "ser");
  for (const ser of serEls) {
    const series = kind === "scatter" || kind === "bubble"
      ? parseXYSeries(ser, kind)
      : parseCategorySeries(ser);
    if (series) {
      series.color = parseSeriesColor(ser, theme);
      spec.series.push(series);
    }
  }

  return spec.series.length > 0 ? spec : null;
}

function findChartTypeElement(plotArea: Element): Element | null {
  for (const child of Array.from(plotArea.childNodes)) {
    if (child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI === NS.c && /Chart$/.test(el.localName)) return el;
  }
  return null;
}

function parseTitle(chart: Element): string | undefined {
  const title = firstChildNS(chart, NS.c, "title");
  if (!title) return undefined;
  // Title text lives in a:t runs under c:tx > c:rich.
  const texts = title.getElementsByTagNameNS(NS.a, "t");
  const joined = Array.from(texts).map((t) => t.textContent ?? "").join("").trim();
  return joined || undefined;
}

function parseCategorySeries(ser: Element): ChartSeries | null {
  const name = parseSeriesName(ser);
  const cat = firstChildNS(ser, NS.c, "cat");
  const val = firstChildNS(ser, NS.c, "val");
  const categories = cat ? cachePoints(cat) : [];
  const values = (val ? cachePoints(val) : []).map(toNum);
  if (values.length === 0) return null;
  return { name, categories, values };
}

function parseXYSeries(ser: Element, kind: ChartKind): ChartSeries | null {
  const name = parseSeriesName(ser);
  const xVal = firstChildNS(ser, NS.c, "xVal");
  const yVal = firstChildNS(ser, NS.c, "yVal");
  const bubble = firstChildNS(ser, NS.c, "bubbleSize");
  const xs = (xVal ? cachePoints(xVal) : []).map(toNum);
  const ys = (yVal ? cachePoints(yVal) : []).map(toNum);
  const rs = (bubble ? cachePoints(bubble) : []).map(toNum);
  const n = Math.min(xs.length, ys.length);
  if (n === 0) return null;
  const points = [];
  for (let i = 0; i < n; i++) {
    points.push(kind === "bubble" ? { x: xs[i], y: ys[i], r: rs[i] ?? 5 } : { x: xs[i], y: ys[i] });
  }
  return { name, categories: [], values: [], points };
}

function parseSeriesName(ser: Element): string {
  const tx = firstChildNS(ser, NS.c, "tx");
  if (tx) {
    const pts = cachePoints(tx);
    if (pts.length > 0) return pts[0];
    const v = firstChildNS(tx, NS.c, "v");
    if (v?.textContent) return v.textContent;
  }
  return "";
}

// Collect cached <c:pt><c:v> values (from strCache/numCache) ordered by idx.
function cachePoints(ref: Element): string[] {
  const pts = ref.getElementsByTagNameNS(NS.c, "pt");
  const byIdx: string[] = [];
  let maxIdx = -1;
  for (const pt of Array.from(pts)) {
    const idx = parseInt(pt.getAttribute("idx") ?? "-1", 10);
    const v = firstChildNS(pt, NS.c, "v");
    if (idx >= 0) {
      byIdx[idx] = v?.textContent ?? "";
      if (idx > maxIdx) maxIdx = idx;
    }
  }
  const out: string[] = [];
  for (let i = 0; i <= maxIdx; i++) out.push(byIdx[i] ?? "");
  return out;
}

function parseSeriesColor(ser: Element, theme?: readonly string[]): string | undefined {
  const spPr = firstChildNS(ser, NS.c, "spPr");
  if (!spPr) return undefined;
  const solidFill = firstChildNS(spPr, NS.a, "solidFill");
  if (!solidFill) return undefined;
  const srgb = firstChildNS(solidFill, NS.a, "srgbClr");
  if (srgb) {
    const val = srgb.getAttribute("val");
    if (val && /^[0-9A-Fa-f]{6}$/.test(val)) return `#${val.toUpperCase()}`;
  }
  const scheme = firstChildNS(solidFill, NS.a, "schemeClr");
  if (scheme && theme) {
    const idx = SCHEME_TO_THEME_INDEX[scheme.getAttribute("val") ?? ""];
    if (idx !== undefined && theme[idx]) return `#${theme[idx]}`;
  }
  return undefined;
}

function childrenNS(parent: Element, ns: string, local: string): Element[] {
  const out: Element[] = [];
  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI === ns && el.localName === local) out.push(el);
  }
  return out;
}

function attrVal(el: Element | null): string | null {
  return el?.getAttribute("val") ?? null;
}

function toNum(s: string): number {
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}
