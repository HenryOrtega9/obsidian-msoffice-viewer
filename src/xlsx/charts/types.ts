import type { AnchorPoint } from "../geometry";

export type ChartKind =
  | "bar"
  | "line"
  | "pie"
  | "doughnut"
  | "scatter"
  | "area"
  | "radar"
  | "bubble"
  | "unsupported";

export interface ChartPoint {
  x: number | null;
  y: number | null;
  r?: number;
}

export interface ChartSeries {
  name: string;
  categories: string[];
  // null = blank/non-numeric point (rendered as a gap, not 0).
  values: (number | null)[];
  points?: ChartPoint[]; // for scatter / bubble
  color?: string;
  // Per-series chart kind, set for combo charts so the renderer can emit a
  // mixed dataset (e.g. a line over bars). Matches the primary kind otherwise.
  kind?: ChartKind;
}

export interface ChartSpec {
  kind: ChartKind;
  rawType?: string; // OOXML element name when unsupported
  barHorizontal?: boolean;
  stacked?: boolean;
  title?: string;
  scatterStyle?: string; // c:scatterStyle val: marker|line|lineMarker|smooth|smoothMarker
  series: ChartSeries[];
}

export interface ChartPlacement {
  from: AnchorPoint;
  to: AnchorPoint | null;
  // Pixel size for a oneCellAnchor chart (null for twoCellAnchor, where `to`
  // sizes the box). Without it oneCellAnchor charts collapsed to a 1x1px box.
  ext: { width: number; height: number } | null;
  // absoluteAnchor box in px (independent of cells); overrides from/to/ext.
  abs: { x: number; y: number; width: number; height: number } | null;
  spec: ChartSpec;
}
