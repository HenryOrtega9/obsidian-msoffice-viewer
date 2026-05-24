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
  x: number;
  y: number;
  r?: number;
}

export interface ChartSeries {
  name: string;
  categories: string[];
  values: number[];
  points?: ChartPoint[]; // for scatter / bubble
  color?: string;
}

export interface ChartSpec {
  kind: ChartKind;
  rawType?: string; // OOXML element name when unsupported
  barHorizontal?: boolean;
  stacked?: boolean;
  title?: string;
  series: ChartSeries[];
}

export interface ChartPlacement {
  from: AnchorPoint;
  to: AnchorPoint | null;
  spec: ChartSpec;
}
