import type JSZip from "jszip";
import type { Chart } from "chart.js";
import { NS, directChild, relId, readXml } from "../ooxml";
import type { Box } from "../geometry";
import type { PptxTheme } from "../themes";
import { warn } from "../warn";
import { parseChartXml } from "../../xlsx/charts/chartXml";
import { renderChartIntoHolder } from "../../xlsx/charts/render";

// Render a chart embedded in a p:graphicFrame. `graphicData` is the
// a:graphicData element whose uri ends in ".../chart"; it wraps a c:chart with
// an r:id pointing at the chart part via the slide's rels. The chart XML format
// is identical to xlsx, so we reuse parseChartXml + the Chart.js core.
export async function renderChartFrame(
  graphicData: Element,
  box: Box,
  zip: JSZip,
  rels: Map<string, string>,
  theme: PptxTheme | null,
  parent: HTMLElement,
): Promise<Chart | null> {
  const chartEl = directChild(graphicData, NS.c, "chart");
  const rid = relId(chartEl);
  const chartPath = rid ? rels.get(rid) : undefined;
  if (!chartPath) return null;

  try {
    const doc = await readXml(zip, chartPath);
    if (!doc) return null;
    const spec = parseChartXml(doc, themeColorArray(theme));
    if (!spec) return null;

    const holder = parent.createDiv({ cls: "docx-claude-pptx-chart-holder" });
    holder.style.left = `${box.left}px`;
    holder.style.top = `${box.top}px`;
    holder.style.width = `${box.width}px`;
    holder.style.height = `${box.height}px`;

    // No LibreOffice escape hatch here: the native renderer is the pptx
    // fallback tier, so LibreOffice has already failed or is absent.
    return renderChartIntoHolder(holder, spec, {});
  } catch (e) {
    warn("chart-frame", e, { chartPath });
    return null;
  }
}

// Flatten the named theme into the index-ordered array parseChartXml expects
// (SCHEME_TO_THEME_INDEX: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink).
function themeColorArray(theme: PptxTheme | null): string[] | undefined {
  if (!theme) return undefined;
  const s = theme.scheme;
  return [
    s.dk1, s.lt1, s.dk2, s.lt2,
    s.accent1, s.accent2, s.accent3, s.accent4, s.accent5, s.accent6,
    s.hlink, s.folHlink,
  ].map((h) => h ?? "");
}
