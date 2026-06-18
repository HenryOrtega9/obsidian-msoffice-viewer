import JSZip from "jszip";
import { warn } from "../warn";
import { ChartPlacement } from "./types";
import { NS, dirOf, readRels, readXml, relId, resolvePath } from "./ooxml";
import { parseDrawingForCharts } from "./drawingXml";
import { parseChartXml } from "./chartXml";

export type { ChartPlacement } from "./types";

// Build a map of worksheet name → embedded chart placements by walking the
// OOXML relationship chain: workbook → sheet → drawing → chart.
export async function loadWorkbookCharts(
  buf: ArrayBuffer,
  theme?: readonly string[],
): Promise<Map<string, ChartPlacement[]>> {
  const result = new Map<string, ChartPlacement[]>();
  try {
    const zip = await JSZip.loadAsync(buf);
    // Quick bail: no chart parts means nothing to do.
    const hasCharts = Object.keys(zip.files).some((p) => /^xl\/charts\/chart\d+\.xml$/.test(p));
    if (!hasCharts) return result;

    const wbDoc = await readXml(zip, "xl/workbook.xml");
    if (!wbDoc) return result;
    const wbRels = await readRels(zip, "xl/_rels/workbook.xml.rels", "xl");

    const sheetEls = wbDoc.getElementsByTagNameNS(NS.ssml, "sheet");
    for (const sheetEl of Array.from(sheetEls)) {
      const name = sheetEl.getAttribute("name");
      const rid = relId(sheetEl);
      if (!name || !rid) continue;
      const sheetPath = wbRels.get(rid);
      if (!sheetPath) continue;

      const placements = await chartsForSheet(zip, sheetPath, theme);
      if (placements.length > 0) result.set(name, placements);
    }
  } catch (e) {
    warn("charts-load", e);
  }
  return result;
}

async function chartsForSheet(
  zip: JSZip,
  sheetPath: string,
  theme?: readonly string[],
): Promise<ChartPlacement[]> {
  const placements: ChartPlacement[] = [];
  const sheetDoc = await readXml(zip, sheetPath);
  if (!sheetDoc) return placements;

  const drawingEls = sheetDoc.getElementsByTagNameNS(NS.ssml, "drawing");
  if (drawingEls.length === 0) return placements;

  const sheetDir = dirOf(sheetPath);
  const sheetRels = await readRels(
    zip,
    `${sheetDir}/_rels/${baseName(sheetPath)}.rels`,
    sheetDir,
  );

  for (const drawingEl of Array.from(drawingEls)) {
    const drawingRid = relId(drawingEl);
    if (!drawingRid) continue;
    const drawingPath = sheetRels.get(drawingRid);
    if (!drawingPath) continue;

    const drawingDoc = await readXml(zip, drawingPath);
    if (!drawingDoc) continue;
    const anchors = parseDrawingForCharts(drawingDoc);
    if (anchors.length === 0) continue;

    const drawingDir = dirOf(drawingPath);
    const drawingRels = await readRels(
      zip,
      `${drawingDir}/_rels/${baseName(drawingPath)}.rels`,
      drawingDir,
    );

    for (const anchor of anchors) {
      const chartPath = drawingRels.get(anchor.chartRelId);
      if (!chartPath) continue;
      const chartDoc = await readXml(zip, chartPath);
      if (!chartDoc) continue;
      const spec = parseChartXml(chartDoc, theme);
      if (!spec) continue;
      placements.push({ from: anchor.from, to: anchor.to, ext: anchor.ext, spec });
    }
  }
  return placements;
}

function baseName(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? path : path.slice(idx + 1);
}

// Re-exported for callers that want to resolve relative targets themselves.
export { resolvePath };
