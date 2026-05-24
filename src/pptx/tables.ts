import { NS, directChild, childrenNS, intAttr } from "./ooxml";
import type { Box } from "./geometry";
import { parseTxBody, renderParagraphsInto, type TextResolveCtx } from "./text";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";

const TABLE_URI_FRAGMENT = "/table";

// Render a p:graphicFrame's a:tbl as an absolutely-positioned HTML table.
// `scale` is px-per-EMU for column/row sizing and cell text. Returns false if
// the graphicFrame does not contain a table.
export function renderTableInto(
  graphicFrame: Element,
  box: Box,
  ctx: TextResolveCtx,
  scale: number,
  parent: HTMLElement,
): boolean {
  const tbl = findTbl(graphicFrame);
  if (!tbl) return false;

  const holder = parent.createDiv({ cls: "docx-claude-pptx-table-holder" });
  holder.style.left = `${box.left}px`;
  holder.style.top = `${box.top}px`;
  holder.style.width = `${box.width}px`;
  holder.style.height = `${box.height}px`;

  const table = holder.createEl("table", { cls: "docx-claude-pptx-table" });

  const grid = directChild(tbl, NS.a, "tblGrid");
  if (grid) {
    const colgroup = table.createEl("colgroup");
    for (const col of childrenNS(grid, NS.a, "gridCol")) {
      const colEl = colgroup.createEl("col");
      const w = intAttr(col, "w", 0) * scale;
      if (w > 0) colEl.style.width = `${w}px`;
    }
  }

  for (const tr of childrenNS(tbl, NS.a, "tr")) {
    const rowEl = table.createEl("tr");
    const h = intAttr(tr, "h", 0) * scale;
    if (h > 0) rowEl.style.height = `${h}px`;
    for (const tc of childrenNS(tr, NS.a, "tc")) {
      // Continuation cells of a horizontal/vertical merge are not rendered;
      // the origin cell carries gridSpan/rowSpan instead.
      if (tc.getAttribute("hMerge") === "1" || tc.getAttribute("vMerge") === "1") continue;
      const td = rowEl.createEl("td");
      const gridSpan = intAttr(tc, "gridSpan", 1);
      const rowSpan = intAttr(tc, "rowSpan", 1);
      if (gridSpan > 1) td.colSpan = gridSpan;
      if (rowSpan > 1) td.rowSpan = rowSpan;
      const tcPr = directChild(tc, NS.a, "tcPr");
      if (tcPr) applyCellStyle(td, tcPr, ctx, scale);
      const txBody = directChild(tc, NS.a, "txBody");
      if (txBody) renderParagraphsInto(parseTxBody(txBody, ctx), td, scale);
    }
  }
  return true;
}

function findTbl(graphicFrame: Element): Element | null {
  const graphic = directChild(graphicFrame, NS.a, "graphic");
  const graphicData = graphic ? directChild(graphic, NS.a, "graphicData") : null;
  if (!graphicData) return null;
  if (!(graphicData.getAttribute("uri") ?? "").includes(TABLE_URI_FRAGMENT)) return null;
  return directChild(graphicData, NS.a, "tbl");
}

function applyCellStyle(td: HTMLElement, tcPr: Element, ctx: TextResolveCtx, scale: number): void {
  const fill = directChild(tcPr, NS.a, "solidFill");
  if (fill) {
    const css = resolveDrawingMlColor(parseColorChoice(fill), ctx.theme, ctx.clrMap);
    if (css) td.style.background = css;
  }
  applyBorder(td, directChild(tcPr, NS.a, "lnL"), "border-left", ctx, scale);
  applyBorder(td, directChild(tcPr, NS.a, "lnR"), "border-right", ctx, scale);
  applyBorder(td, directChild(tcPr, NS.a, "lnT"), "border-top", ctx, scale);
  applyBorder(td, directChild(tcPr, NS.a, "lnB"), "border-bottom", ctx, scale);

  const anchor = tcPr.getAttribute("anchor"); // t / ctr / b
  td.style.verticalAlign = anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top";
}

function applyBorder(
  td: HTMLElement,
  ln: Element | null,
  cssProp: string,
  ctx: TextResolveCtx,
  scale: number,
): void {
  if (!ln) return;
  if (directChild(ln, NS.a, "noFill")) {
    td.style.setProperty(cssProp, "none");
    return;
  }
  const solid = directChild(ln, NS.a, "solidFill");
  const color = solid ? resolveDrawingMlColor(parseColorChoice(solid), ctx.theme, ctx.clrMap) : null;
  const w = intAttr(ln, "w", 0);
  if (!color && w === 0) return;
  const px = w > 0 ? Math.max(0.5, w * scale) : 1;
  td.style.setProperty(cssProp, `${px}px solid ${color ?? "#000000"}`);
}
