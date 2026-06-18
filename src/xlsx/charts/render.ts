import {
  ArcElement,
  BarController,
  BarElement,
  BubbleController,
  CategoryScale,
  Chart,
  type ChartConfiguration,
  DoughnutController,
  Filler,
  Legend,
  LineController,
  LineElement,
  LinearScale,
  PieController,
  PointElement,
  RadarController,
  RadialLinearScale,
  ScatterController,
  Title,
  Tooltip,
} from "chart.js";
import type { GridContext } from "../grid";
import { anchorRangeToBox } from "../geometry";
import { warn } from "../warn";
import { ChartKind, ChartPlacement, ChartSeries, ChartSpec } from "./types";

const MAX_CHARTS_PER_SHEET = 10;

export const PALETTE = [
  "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
  "#264478", "#9E480E", "#636363", "#997300", "#255E91", "#43682B",
];

let registered = false;
export function ensureRegistered(): void {
  if (registered) return;
  Chart.register(
    BarController, LineController, PieController, DoughnutController,
    ScatterController, RadarController, BubbleController,
    CategoryScale, LinearScale, RadialLinearScale,
    BarElement, LineElement, PointElement, ArcElement,
    Filler, Tooltip, Legend, Title,
  );
  registered = true;
}

export interface RenderChartsOpts {
  onRequestLibreOffice?: () => void;
}

// Render embedded charts into an overlay layer above the grid. Returns the
// live Chart instances so the caller can destroy them on unload / sheet switch.
export function renderSheetCharts(
  placements: ChartPlacement[],
  ctx: GridContext,
  opts: RenderChartsOpts = {},
): Chart[] {
  if (placements.length === 0) return [];
  ensureRegistered();

  const layer = ctx.sheetWrapEl.createDiv({ cls: "docx-claude-xlsx-chart-layer" });
  const charts: Chart[] = [];

  let count = 0;
  for (const placement of placements) {
    if (count >= MAX_CHARTS_PER_SHEET) {
      warn("charts-cap", null, { cap: MAX_CHARTS_PER_SHEET, total: placements.length });
      break;
    }
    count++;

    const box = anchorRangeToBox(placement.from, placement.to, placement.ext, ctx);
    const holder = layer.createDiv({ cls: "docx-claude-xlsx-chart-holder" });
    holder.style.left = `${box.left}px`;
    holder.style.top = `${box.top}px`;
    holder.style.width = `${box.width}px`;
    holder.style.height = `${box.height}px`;

    const chart = renderChartIntoHolder(holder, placement.spec, opts);
    if (chart) charts.push(chart);
  }

  return charts;
}

// Placement-agnostic core: register controllers, build the config, instantiate
// Chart.js into `holder`, falling back to a placeholder on unsupported types or
// errors. Shared with the pptx renderer so charts behave identically and no
// second copy of buildConfig / Chart.js registration ships in the bundle.
export function renderChartIntoHolder(
  holder: HTMLElement,
  spec: ChartSpec,
  opts: RenderChartsOpts = {},
): Chart | null {
  ensureRegistered();
  const config = spec.kind === "unsupported" ? null : buildConfig(spec);
  if (!config) {
    renderPlaceholder(holder, spec, opts);
    return null;
  }
  try {
    const canvas = holder.createEl("canvas");
    return new Chart(canvas, config);
  } catch (e) {
    warn("chart-render", e, { kind: spec.kind });
    holder.empty();
    renderPlaceholder(holder, spec, opts);
    return null;
  }
}

function renderPlaceholder(holder: HTMLElement, spec: ChartSpec, opts: RenderChartsOpts): void {
  holder.addClass("docx-claude-xlsx-chart-placeholder");
  const label = spec.rawType ? `Unsupported chart: ${spec.rawType}` : "Chart could not be rendered";
  holder.createDiv({ cls: "docx-claude-xlsx-chart-ph-label", text: label });
  if (opts.onRequestLibreOffice) {
    const btn = holder.createEl("button", {
      cls: "docx-claude-xlsx-chart-lo-btn",
      text: "Render this sheet with LibreOffice",
    });
    btn.addEventListener("click", (ev) => {
      ev.preventDefault();
      opts.onRequestLibreOffice?.();
    });
  }
}

export function buildConfig(spec: ChartSpec): ChartConfiguration | null {
  const cjsType = mapType(spec.kind);
  if (!cjsType) return null;

  const titlePlugin = spec.title
    ? { title: { display: true, text: spec.title } }
    : {};
  const common = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: spec.series.length > 1 || isPieLike(spec.kind) }, ...titlePlugin },
  };

  if (spec.kind === "scatter" || spec.kind === "bubble") {
    return {
      type: cjsType,
      data: {
        datasets: spec.series.map((s, i) => ({
          label: s.name || `Series ${i + 1}`,
          data: (s.points ?? []).map((p) => ({ x: p.x, y: p.y, ...(p.r != null ? { r: p.r } : {}) })),
          backgroundColor: s.color ?? PALETTE[i % PALETTE.length],
        })),
      },
      options: common,
    } as ChartConfiguration;
  }

  const rawLabels = spec.series[0]?.categories ?? [];

  if (isPieLike(spec.kind)) {
    const s = spec.series[0];
    if (!s) return null;
    return {
      type: cjsType,
      data: {
        labels: rawLabels,
        datasets: [{
          data: s.values,
          backgroundColor: s.values.map((_, i) => PALETTE[i % PALETTE.length]),
        }],
      },
      options: common,
    } as ChartConfiguration;
  }

  // Synthesize Excel-style 1..n category labels when <c:cat> is absent, and pad
  // when series[0]'s categories are shorter than the longest series — otherwise
  // chart.js draws 0 of N points (a completely blank chart).
  const n = Math.max(0, ...spec.series.map((s) => s.values.length));
  const labels = Array.from({ length: n }, (_, i) => rawLabels[i] ?? String(i + 1));

  const stacked = spec.stacked === true;
  const fill = spec.kind === "area";
  return {
    type: cjsType,
    data: {
      labels,
      datasets: spec.series.map((s, i) => seriesDataset(s, i, fill)),
    },
    options: {
      ...common,
      indexAxis: spec.barHorizontal ? "y" : "x",
      scales: {
        x: { stacked },
        y: { stacked },
      },
    },
  } as unknown as ChartConfiguration;
}

function seriesDataset(s: ChartSeries, i: number, fill: boolean): Record<string, unknown> {
  const color = s.color ?? PALETTE[i % PALETTE.length];
  return {
    label: s.name || `Series ${i + 1}`,
    data: s.values,
    backgroundColor: fill ? hexWithAlpha(color, 0.35) : color,
    borderColor: color,
    fill,
  };
}

function mapType(kind: ChartKind): ChartConfiguration["type"] | null {
  switch (kind) {
    case "bar": return "bar";
    case "line": return "line";
    case "area": return "line";
    case "pie": return "pie";
    case "doughnut": return "doughnut";
    case "scatter": return "scatter";
    case "bubble": return "bubble";
    case "radar": return "radar";
    default: return null;
  }
}

function isPieLike(kind: ChartKind): boolean {
  return kind === "pie" || kind === "doughnut";
}

function hexWithAlpha(hex: string, alpha: number): string {
  if (!hex.startsWith("#") || hex.length < 7) return hex;
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}
