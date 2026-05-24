import type ExcelJS from "exceljs";
import type { GridContext } from "../grid";
import type { MergeRect } from "../merges";
import { warn } from "../warn";
import { parseCfRanges } from "./values";
import { applyColorScale, applyDataBar, applyIconSet } from "./visuals";
import {
  applyAboveAverage,
  applyCellIs,
  applyContainsText,
  applyTimePeriod,
  applyTop10,
} from "./predicates";

interface CfRule {
  type?: string;
  priority?: number;
  stopIfTrue?: boolean;
  [key: string]: unknown;
}

interface CfEntry {
  ref?: string;
  rules?: CfRule[];
}

interface FlatRule {
  ranges: MergeRect[];
  rule: CfRule;
}

// Apply all conditional-formatting rules for a worksheet onto the rendered
// grid. Rules are sorted by priority (lower = higher precedence); predicate
// rules honor stopIfTrue so a higher-priority match blocks lower ones.
export function applyConditionalFormatting(
  ws: ExcelJS.Worksheet,
  ctx: GridContext,
  theme?: readonly string[],
): void {
  const cfs = (ws as unknown as { conditionalFormattings?: CfEntry[] }).conditionalFormattings;
  if (!Array.isArray(cfs) || cfs.length === 0) return;

  const flat: FlatRule[] = [];
  for (const cf of cfs) {
    if (!cf.ref || !Array.isArray(cf.rules)) continue;
    const ranges = parseCfRanges(cf.ref);
    if (ranges.length === 0) continue;
    for (const rule of cf.rules) flat.push({ ranges, rule });
  }

  flat.sort((a, b) => (a.rule.priority ?? 1e9) - (b.rule.priority ?? 1e9));

  const stopped = new Set<string>();

  for (const { ranges, rule } of flat) {
    const stopIfTrue = rule.stopIfTrue === true;
    try {
      switch (rule.type) {
        case "colorScale":
          applyColorScale(ws, ctx, ranges, rule as never, theme);
          break;
        case "dataBar":
          applyDataBar(ws, ctx, ranges, rule as never, theme);
          break;
        case "iconSet":
          applyIconSet(ws, ctx, ranges, rule as never);
          break;
        case "cellIs":
          applyCellIs(ws, ctx, ranges, rule as never, stopped, stopIfTrue, theme);
          break;
        case "top10":
          applyTop10(ws, ctx, ranges, rule as never, stopped, stopIfTrue, theme);
          break;
        case "aboveAverage":
          applyAboveAverage(ws, ctx, ranges, rule as never, stopped, stopIfTrue, theme);
          break;
        case "containsText":
          applyContainsText(ws, ctx, ranges, rule as never, stopped, stopIfTrue, theme);
          break;
        case "timePeriod":
          applyTimePeriod(ws, ctx, ranges, rule as never, stopped, stopIfTrue, theme);
          break;
        case "expression":
          warn("cf-expression", "formula evaluation unsupported", { ref: ranges });
          break;
        default:
          warn("cf-unknown", null, { type: rule.type });
      }
    } catch (e) {
      warn("cf-apply", e, { type: rule.type });
    }
  }
}
