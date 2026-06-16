import type ExcelJS from "exceljs";
import type { MergeRect } from "../merges";
import { ExcelColorRef, resolveExcelColor } from "../colors";
import type { GridContext } from "../grid";
import {
  Cfvo,
  cfvoToNumber,
  collectNumericValues,
  forEachCellInRanges,
  numericCellValue,
} from "./values";

type RGB = [number, number, number];

function colorToRgb(color: ExcelColorRef | undefined, theme?: readonly string[]): RGB | null {
  const css = resolveExcelColor(color, theme);
  if (!css || !css.startsWith("#")) return null;
  const hex = css.slice(1);
  if (hex.length < 6) return null;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  if ([r, g, b].some((n) => Number.isNaN(n))) return null;
  return [r, g, b];
}

function rgbToCss([r, g, b]: RGB): string {
  return `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`;
}

function lerp(a: number, b: number, t: number): number {
  return a + (b - a) * t;
}

// --- Color scale -----------------------------------------------------------

export function applyColorScale(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { cfvo?: Cfvo[]; color?: ExcelColorRef[] },
  theme?: readonly string[],
): void {
  const cfvo = rule.cfvo ?? [];
  const colors = rule.color ?? [];
  if (cfvo.length < 2 || colors.length < 2) return;

  const values = collectNumericValues(ws, ranges);
  if (values.length === 0) return;
  const sorted = [...values].sort((a, b) => a - b);
  const min = sorted[0];
  const max = sorted[sorted.length - 1];

  const stops = cfvo.map((c) => cfvoToNumber(c, sorted, min, max));
  const rgbStops = colors.map((c) => colorToRgb(c, theme));
  if (rgbStops.some((c) => c === null)) return;

  forEachCellInRanges(ranges, (r, c) => {
    const v = numericCellValue(ws, r, c);
    if (v == null) return;
    const td = ctx.cellMap.get(`${r}:${c}`);
    if (!td) return;
    const color = interpolateScale(v, stops, rgbStops as RGB[]);
    td.style.backgroundColor = rgbToCss(color);
  });
}

function interpolateScale(v: number, stops: number[], colors: RGB[]): RGB {
  if (v <= stops[0]) return colors[0];
  const last = stops.length - 1;
  if (v >= stops[last]) return colors[last];
  for (let i = 0; i < last; i++) {
    const lo = stops[i];
    const hi = stops[i + 1];
    if (v >= lo && v <= hi) {
      const t = hi === lo ? 0 : (v - lo) / (hi - lo);
      return [
        lerp(colors[i][0], colors[i + 1][0], t),
        lerp(colors[i][1], colors[i + 1][1], t),
        lerp(colors[i][2], colors[i + 1][2], t),
      ];
    }
  }
  return colors[last];
}

// --- Data bar --------------------------------------------------------------

export function applyDataBar(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: {
    cfvo?: Cfvo[];
    color?: ExcelColorRef;
    gradient?: boolean;
    minLength?: number;
    maxLength?: number;
    axisPosition?: string;
  },
  theme?: readonly string[],
): void {
  const cfvo = rule.cfvo ?? [];
  const values = collectNumericValues(ws, ranges);
  if (values.length === 0) return;
  const sorted = [...values].sort((a, b) => a - b);
  const dataMin = sorted[0];
  const dataMax = sorted[sorted.length - 1];

  let min = cfvo[0] ? cfvoToNumber(cfvo[0], sorted, dataMin, dataMax) : dataMin;
  let max = cfvo[1] ? cfvoToNumber(cfvo[1], sorted, dataMin, dataMax) : dataMax;
  if (max === min) max = min + 1;

  const barRgb = colorToRgb(rule.color, theme) ?? [99, 142, 198];
  const barColor = rgbToCss(barRgb);
  const lighter = rgbToCss([
    Math.min(255, barRgb[0] + 60),
    Math.min(255, barRgb[1] + 60),
    Math.min(255, barRgb[2] + 60),
  ]);
  const maxLen = rule.maxLength ?? 100;

  const hasNegative = min < 0;
  const axisPct = hasNegative ? (-min / (max - min)) * 100 : 0;

  forEachCellInRanges(ranges, (r, c) => {
    const v = numericCellValue(ws, r, c);
    if (v == null) return;
    const td = ctx.cellMap.get(`${r}:${c}`);
    if (!td) return;

    const frac = Math.max(0, Math.min(1, (v - min) / (max - min)));

    if (!hasNegative) {
      const pct = frac * maxLen;
      const end = rule.gradient ? lighter : barColor;
      td.style.backgroundImage = `linear-gradient(to right, ${barColor} 0%, ${end} ${pct}%, transparent ${pct}%)`;
    } else {
      // Negative-aware: bar grows from the zero axis.
      const valuePct = frac * 100;
      if (v >= 0) {
        td.style.backgroundImage = `linear-gradient(to right, transparent ${axisPct}%, ${barColor} ${axisPct}%, ${barColor} ${valuePct}%, transparent ${valuePct}%)`;
      } else {
        td.style.backgroundImage = `linear-gradient(to right, transparent ${valuePct}%, #d65a5a ${valuePct}%, #d65a5a ${axisPct}%, transparent ${axisPct}%)`;
      }
    }
    td.style.backgroundClip = "padding-box";
  });
}

// --- Icon set --------------------------------------------------------------

interface IconGlyph {
  glyph: string;
  color: string;
}

// Pick an icon family + colors by name. Counts (3/4/5) come from the leading
// digit. Traffic lights / signs use colored emoji; everything else uses a
// monochrome glyph tinted along a red→green ramp.
function iconsFor(name: string): IconGlyph[] {
  const count = parseInt(name.charAt(0), 10) || 3;
  const ramp = colorRamp(count);

  if (/TrafficLights|Signs/.test(name)) {
    const lights = ["\u{1F534}", "\u{1F7E1}", "\u{1F7E2}"]; // 🔴 🟡 🟢
    if (count === 3) return lights.map((g) => ({ glyph: g, color: "" }));
  }
  if (/Arrows/.test(name)) {
    const arrows = count >= 5
      ? ["↓", "↘", "→", "↗", "↑"]
      : count === 4
        ? ["↓", "↘", "↗", "↑"]
        : ["↓", "→", "↑"];
    return arrows.map((g, i) => ({ glyph: g, color: ramp[i] }));
  }
  if (/Symbols/.test(name)) {
    const syms = ["✖", "!", "✔"]; // ✖ ! ✔
    if (count === 3) return syms.map((g, i) => ({ glyph: g, color: ramp[i] }));
  }
  if (/Flags/.test(name)) {
    return Array.from({ length: count }, (_, i) => ({ glyph: "⚑", color: ramp[i] }));
  }
  if (/Stars|Rating|Quarters|Boxes/.test(name)) {
    return Array.from({ length: count }, (_, i) => ({ glyph: "●", color: ramp[i] }));
  }
  // Default: colored dots low→high.
  return Array.from({ length: count }, (_, i) => ({ glyph: "●", color: ramp[i] }));
}

function colorRamp(count: number): string[] {
  // red → amber → green ramp
  const stops: RGB[] = [
    [216, 90, 90],
    [230, 175, 70],
    [112, 173, 71],
  ];
  const out: string[] = [];
  for (let i = 0; i < count; i++) {
    const t = count === 1 ? 0 : i / (count - 1);
    const seg = t * (stops.length - 1);
    const lo = Math.floor(seg);
    const hi = Math.min(stops.length - 1, lo + 1);
    const f = seg - lo;
    out.push(rgbToCss([
      lerp(stops[lo][0], stops[hi][0], f),
      lerp(stops[lo][1], stops[hi][1], f),
      lerp(stops[lo][2], stops[hi][2], f),
    ]));
  }
  return out;
}

export function applyIconSet(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  ranges: MergeRect[],
  rule: { cfvo?: Cfvo[]; iconSet?: string; reverse?: boolean; showValue?: boolean },
): void {
  const name = rule.iconSet ?? "3TrafficLights1";
  if (name === "NoIcons") return;
  const icons = iconsFor(name);
  if (icons.length === 0) return;

  const cfvo = rule.cfvo ?? [];
  const values = collectNumericValues(ws, ranges);
  if (values.length === 0) return;
  const sorted = [...values].sort((a, b) => a - b);
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  const thresholds = cfvo.map((c) => cfvoToNumber(c, sorted, min, max));

  forEachCellInRanges(ranges, (r, c) => {
    const v = numericCellValue(ws, r, c);
    if (v == null) return;
    const td = ctx.cellMap.get(`${r}:${c}`);
    if (!td) return;

    let bucket = 0;
    for (let i = 0; i < thresholds.length; i++) {
      if (v >= thresholds[i]) bucket = i;
    }
    bucket = Math.min(bucket, icons.length - 1);
    const idx = rule.reverse ? icons.length - 1 - bucket : bucket;
    const icon = icons[idx];

    const span = document.createElement("span");
    span.className = "docx-claude-xlsx-cf-icon";
    span.textContent = icon.glyph;
    if (icon.color) span.style.color = icon.color;
    td.prepend(span);

    if (rule.showValue === false) {
      // Hide the underlying text but keep the icon.
      for (const node of Array.from(td.childNodes)) {
        if (node !== span && node.nodeType === 3) node.textContent = "";
      }
    }
  });
}
