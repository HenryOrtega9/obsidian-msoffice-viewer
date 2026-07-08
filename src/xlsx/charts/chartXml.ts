import { ChartKind, ChartSeries, ChartSpec } from "./types";
import { NS, firstChildNS } from "./ooxml";
import { applyColorMods } from "../../pptx/colors";
import { format as numfmtFormat } from "numfmt";
import { warn } from "../warn";

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

// The theme array from xlsx/themes.ts is in Excel index order: 0=lt1, 1=dk1,
// 2=lt2, 3=dk2, then accents. tx maps to dark, bg to light.
const SCHEME_TO_THEME_INDEX: Record<string, number> = {
  lt1: 0, bg1: 0,
  dk1: 1, tx1: 1,
  lt2: 2, bg2: 2,
  dk2: 3, tx2: 3,
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

  // A plotArea can carry several *Chart elements (combo charts, e.g. a bar chart
  // plus a line chart). The primary (first supported) sets the base kind and
  // bar/grouping options; every supported element contributes its series tagged
  // with its own kind, so the line-over-bar series isn't dropped.
  const typeEls = findChartTypeElements(plotArea);
  if (typeEls.length === 0) return null;
  const primaryEl = typeEls.find((el) => TYPE_TO_KIND[el.localName]) ?? typeEls[0];
  const primaryKind = TYPE_TO_KIND[primaryEl.localName] ?? "unsupported";

  const spec: ChartSpec = {
    kind: primaryKind,
    rawType: primaryKind === "unsupported" ? primaryEl.localName : undefined,
    title: parseTitle(chart),
    series: [],
  };

  if (primaryKind === "bar") {
    const barDir = attrVal(firstChildNS(primaryEl, NS.c, "barDir"));
    spec.barHorizontal = barDir === "bar";
  }
  if (primaryKind === "bar" || primaryKind === "line" || primaryKind === "area") {
    const grouping = attrVal(firstChildNS(primaryEl, NS.c, "grouping"));
    spec.stacked = grouping === "stacked" || grouping === "percentStacked";
    spec.percentStacked = grouping === "percentStacked";
  }
  if (primaryKind === "scatter") {
    spec.scatterStyle = attrVal(firstChildNS(primaryEl, NS.c, "scatterStyle")) ?? undefined;
  }

  if (primaryKind === "unsupported") return spec;

  const primaryIsXY = primaryKind === "scatter" || primaryKind === "bubble";
  for (const typeEl of typeEls) {
    const kind = TYPE_TO_KIND[typeEl.localName];
    if (!kind) continue; // skip an unsupported sub-chart within a combo
    const isXY = kind === "scatter" || kind === "bubble";
    // chart.js can't mix a category-axis chart (bar/line/area/radar) with an
    // XY chart (scatter/bubble) on one canvas, so only take series whose axis
    // shape matches the primary's.
    if (isXY !== primaryIsXY) continue;
    for (const ser of childrenNS(typeEl, NS.c, "ser")) {
      const series = isXY ? parseXYSeries(ser, kind) : parseCategorySeries(ser);
      if (!series) continue;
      series.color = parseSeriesColor(ser, kind, theme);
      series.pointColors = parsePointColors(ser, theme);
      series.kind = kind;
      spec.series.push(series);
    }
  }

  return spec.series.length > 0 ? spec : null;
}

function findChartTypeElements(plotArea: Element): Element[] {
  const out: Element[] = [];
  for (const child of Array.from(plotArea.childNodes)) {
    if (child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.namespaceURI === NS.c && /Chart$/.test(el.localName)) out.push(el);
  }
  return out;
}

// First DIRECT child matching ns/local (childNodes scan), unlike firstChildNS
// which is a descendant search and can reach into a nested element.
function firstDirectChildNS(parent: Element, ns: string, local: string): Element | null {
  const list = childrenNS(parent, ns, local);
  return list.length > 0 ? list[0] : null;
}

function parseTitle(chart: Element): string | undefined {
  // DIRECT child of c:chart only — firstChildNS is a descendant search and would
  // otherwise grab a c:title nested inside an axis (c:valAx/c:catAx).
  const title = firstDirectChildNS(chart, NS.c, "title");
  if (!title) return undefined;
  // Title text lives in a:t runs under c:tx > c:rich.
  const texts = title.getElementsByTagNameNS(NS.a, "t");
  const joined = Array.from(texts).map((t) => t.textContent ?? "").join("").trim();
  if (joined) return joined;
  // Fallback: a title bound to a cell reference carries its text in the
  // c:tx > c:strRef/c:numRef cache (no a:t runs).
  const tx = firstChildNS(title, NS.c, "tx");
  if (tx) {
    const cached = cachePoints(tx).join("").trim();
    if (cached) return cached;
  }
  return undefined;
}

function parseCategorySeries(ser: Element): ChartSeries | null {
  const name = parseSeriesName(ser);
  const cat = firstChildNS(ser, NS.c, "cat");
  const val = firstChildNS(ser, NS.c, "val");
  const categories = cat ? formatCategories(cat, cachePoints(cat)) : [];
  const values = (val ? cachePoints(val) : []).map(toNum);
  if (values.length === 0) return null;
  return { name, categories, values };
}

// Date/number category caches carry raw serials plus a c:formatCode; without
// applying it, a date axis shows "45292" instead of "Jan-24".
function formatCategories(catEl: Element, raw: string[]): string[] {
  const fcEl = firstChildNS(catEl, NS.c, "formatCode");
  const fc = fcEl?.textContent?.trim();
  if (!fc || fc === "General" || fc === "@") return raw;
  return raw.map((s) => {
    const n = parseFloat(s);
    if (!Number.isFinite(n)) return s;
    try {
      return numfmtFormat(fc, n);
    } catch (e) {
      warn("chart-cat-fmt", e, { fc, s });
      return s;
    }
  });
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
  const pts = leafLevelScope(ref).getElementsByTagNameNS(NS.c, "pt");
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

// Hierarchical (multiLvlStrRef) categories list one <c:lvl> per level, each
// repeating idx 0..n, so collecting pts across all levels collides. The first
// (leaf) level aligns 1:1 with the value points, so scope to it when present.
function leafLevelScope(ref: Element): Element {
  const cache = firstChildNS(ref, NS.c, "multiLvlStrCache");
  if (cache) {
    const lvl = firstDirectChildNS(cache, NS.c, "lvl");
    if (lvl) return lvl;
  }
  return ref;
}

function parseSeriesColor(
  ser: Element,
  kind: ChartKind,
  theme?: readonly string[],
): string | undefined {
  const spPr = firstChildNS(ser, NS.c, "spPr");
  if (!spPr) return undefined;
  // A line series' visible color is its stroke (spPr > a:ln > solidFill); every
  // other kind uses the body fill (spPr > solidFill). Use DIRECT-child lookups
  // so a descendant search doesn't pick the wrong solidFill (e.g. a marker/line
  // fill for a bar, or a body fill for a line).
  let solidFill: Element | null = null;
  if (kind === "line") {
    const ln = firstDirectChildNS(spPr, NS.a, "ln");
    solidFill = ln ? firstDirectChildNS(ln, NS.a, "solidFill") : null;
    if (!solidFill) solidFill = firstDirectChildNS(spPr, NS.a, "solidFill");
  } else {
    solidFill = firstDirectChildNS(spPr, NS.a, "solidFill");
  }
  return solidFill ? solidFillHex(solidFill, theme) : undefined;
}

function solidFillHex(solidFill: Element, theme?: readonly string[]): string | undefined {
  const srgb = firstDirectChildNS(solidFill, NS.a, "srgbClr");
  if (srgb) {
    const val = srgb.getAttribute("val");
    if (val && /^[0-9A-Fa-f]{6}$/.test(val)) return `#${applyColorMods(val.toUpperCase(), srgb)}`;
  }
  const scheme = firstDirectChildNS(solidFill, NS.a, "schemeClr");
  if (scheme && theme) {
    const idx = SCHEME_TO_THEME_INDEX[scheme.getAttribute("val") ?? ""];
    if (idx !== undefined && theme[idx]) return `#${applyColorMods(theme[idx], scheme)}`;
  }
  return undefined;
}

// Explicit per-point fills (c:dPt > c:spPr > a:solidFill), keyed by c:idx.
// Pie/doughnut slices are colored this way when the user picks slice colors.
function parsePointColors(
  ser: Element,
  theme?: readonly string[],
): (string | undefined)[] | undefined {
  const dPts = childrenNS(ser, NS.c, "dPt");
  if (dPts.length === 0) return undefined;
  const out: (string | undefined)[] = [];
  for (const dPt of dPts) {
    const idxEl = firstDirectChildNS(dPt, NS.c, "idx");
    const idx = parseInt(idxEl?.getAttribute("val") ?? "-1", 10);
    if (idx < 0) continue;
    const spPr = firstDirectChildNS(dPt, NS.c, "spPr");
    const solid = spPr ? firstDirectChildNS(spPr, NS.a, "solidFill") : null;
    if (solid) out[idx] = solidFillHex(solid, theme);
  }
  return out.length > 0 ? out : undefined;
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

// Non-finite/blank/text points become null so chart.js draws a GAP (line/area)
// or skips the point (scatter), rather than a false dip to 0. A genuine "0"
// stays 0 (parseFloat("0") is finite).
function toNum(s: string): number | null {
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}
