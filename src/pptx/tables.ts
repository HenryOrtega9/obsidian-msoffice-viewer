import { NS, directChild, childrenNS, intAttr } from "./ooxml";
import type { Box } from "./geometry";
import { parseTxBody, renderParagraphsInto, type TextResolveCtx } from "./text";
import { parseColorChoice, resolveDrawingMlColor } from "./colors";
import { resolveFillCss } from "./shapes";

const TABLE_URI_FRAGMENT = "/table";

export interface TableRenderCtx extends TextResolveCtx {
  tableStyles?: Document | null; // parsed ppt/tableStyles.xml
}

// Style computed for one part (wholeTbl / firstRow / band1H / ...) of a table
// style definition.
interface PartStyle {
  fillCss: string | null;
  bold: boolean | null;
  colorCss: string | null;
  insideHCss: string | null; // default cell borders from tcBdr
  insideVCss: string | null;
}

// Render a p:graphicFrame's a:tbl as an absolutely-positioned HTML table.
// `scale` is px-per-EMU for column/row sizing and cell text. Returns false if
// the graphicFrame does not contain a table. Cell appearance layers, lowest
// first: referenced table style (wholeTbl -> banding -> first/last row), then
// explicit tcPr fills/borders.
export function renderTableInto(
  graphicFrame: Element,
  box: Box,
  ctx: TableRenderCtx,
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

  const tblPr = directChild(tbl, NS.a, "tblPr");
  const styleParts = resolveTableStyle(tbl, ctx);
  const firstRowOn = tblPr?.getAttribute("firstRow") === "1";
  const bandRowOn = tblPr?.getAttribute("bandRow") === "1";

  const grid = directChild(tbl, NS.a, "tblGrid");
  if (grid) {
    const colgroup = table.createEl("colgroup");
    for (const col of childrenNS(grid, NS.a, "gridCol")) {
      const colEl = colgroup.createEl("col");
      const w = intAttr(col, "w", 0) * scale;
      if (w > 0) colEl.style.width = `${w}px`;
    }
  }

  const rows = childrenNS(tbl, NS.a, "tr");
  rows.forEach((tr, rowIdx) => {
    const rowEl = table.createEl("tr");
    const h = intAttr(tr, "h", 0) * scale;
    if (h > 0) rowEl.style.height = `${h}px`;

    // Pick the style part for this row: header, else banding (band 1 starts on
    // the first non-header row), else the whole-table default.
    let part = styleParts?.wholeTbl ?? null;
    if (styleParts) {
      if (firstRowOn && rowIdx === 0 && styleParts.firstRow) {
        part = mergeParts(styleParts.wholeTbl, styleParts.firstRow);
      } else if (bandRowOn) {
        const bandIdx = firstRowOn ? rowIdx - 1 : rowIdx;
        const band = bandIdx % 2 === 0 ? styleParts.band1H : styleParts.band2H;
        if (band) part = mergeParts(styleParts.wholeTbl, band);
      }
    }

    for (const tc of childrenNS(tr, NS.a, "tc")) {
      // Continuation cells of a horizontal/vertical merge are not rendered;
      // the origin cell carries gridSpan/rowSpan instead.
      if (tc.getAttribute("hMerge") === "1" || tc.getAttribute("vMerge") === "1") continue;
      const td = rowEl.createEl("td");
      const gridSpan = intAttr(tc, "gridSpan", 1);
      const rowSpan = intAttr(tc, "rowSpan", 1);
      if (gridSpan > 1) td.colSpan = gridSpan;
      if (rowSpan > 1) td.rowSpan = rowSpan;

      if (part) {
        if (part.fillCss) td.style.background = part.fillCss;
        if (part.insideHCss) {
          td.style.borderTop = part.insideHCss;
          td.style.borderBottom = part.insideHCss;
        }
        if (part.insideVCss) {
          td.style.borderLeft = part.insideVCss;
          td.style.borderRight = part.insideVCss;
        }
        if (part.bold) td.style.fontWeight = "bold";
        if (part.colorCss) td.style.color = part.colorCss;
      }

      const tcPr = directChild(tc, NS.a, "tcPr");
      if (tcPr) applyCellStyle(td, tcPr, ctx, scale);
      const txBody = directChild(tc, NS.a, "txBody");
      if (txBody) renderParagraphsInto(parseTxBody(txBody, ctx), td, scale);
    }
  });
  return true;
}

interface TableStyleParts {
  wholeTbl: PartStyle | null;
  firstRow: PartStyle | null;
  band1H: PartStyle | null;
  band2H: PartStyle | null;
}

// Look up the a:tblStyle in tableStyles.xml whose styleId matches the table's
// a:tableStyleId, and pre-resolve the row-level parts used here.
function resolveTableStyle(tbl: Element, ctx: TableRenderCtx): TableStyleParts | null {
  const tblPr = directChild(tbl, NS.a, "tblPr");
  const idEl = tblPr ? directChild(tblPr, NS.a, "tableStyleId") : null;
  const styleId = idEl?.textContent?.trim();
  const doc = ctx.tableStyles;
  if (!styleId || !doc) return null;
  const styles = Array.from(doc.getElementsByTagNameNS(NS.a, "tblStyle"));
  const def = styles.find((s) => s.getAttribute("styleId") === styleId);
  if (!def) return null;
  const partOf = (name: string): PartStyle | null => {
    const el = directChild(def, NS.a, name);
    return el ? readPartStyle(el, ctx) : null;
  };
  return {
    wholeTbl: partOf("wholeTbl"),
    firstRow: partOf("firstRow"),
    band1H: partOf("band1H"),
    band2H: partOf("band2H"),
  };
}

function mergeParts(base: PartStyle | null, over: PartStyle): PartStyle {
  return {
    fillCss: over.fillCss ?? base?.fillCss ?? null,
    bold: over.bold ?? base?.bold ?? null,
    colorCss: over.colorCss ?? base?.colorCss ?? null,
    insideHCss: over.insideHCss ?? base?.insideHCss ?? null,
    insideVCss: over.insideVCss ?? base?.insideVCss ?? null,
  };
}

// A table-style part carries a:tcStyle (fill + tcBdr) and a:tcTxStyle
// (b="on"/"off" + a color choice).
function readPartStyle(partEl: Element, ctx: TableRenderCtx): PartStyle {
  const out: PartStyle = { fillCss: null, bold: null, colorCss: null, insideHCss: null, insideVCss: null };
  const tcStyle = directChild(partEl, NS.a, "tcStyle");
  if (tcStyle) {
    const fill = directChild(tcStyle, NS.a, "fill");
    if (fill) out.fillCss = resolveFillCss(fill, ctx);
    const fillRef = directChild(tcStyle, NS.a, "fillRef");
    if (!out.fillCss && fillRef) {
      out.fillCss = resolveDrawingMlColor(parseColorChoice(fillRef), ctx.theme, ctx.clrMap);
    }
    const tcBdr = directChild(tcStyle, NS.a, "tcBdr");
    if (tcBdr) {
      out.insideHCss = borderCssOf(directChild(tcBdr, NS.a, "insideH"), ctx);
      out.insideVCss = borderCssOf(directChild(tcBdr, NS.a, "insideV"), ctx);
    }
  }
  const tcTxStyle = directChild(partEl, NS.a, "tcTxStyle");
  if (tcTxStyle) {
    const b = tcTxStyle.getAttribute("b");
    if (b != null) out.bold = b === "on" || b === "1";
    out.colorCss = resolveDrawingMlColor(parseColorChoice(tcTxStyle), ctx.theme, ctx.clrMap);
  }
  return out;
}

function borderCssOf(side: Element | null, ctx: TableRenderCtx): string | null {
  const ln = side ? directChild(side, NS.a, "ln") : null;
  if (!ln) return null;
  if (directChild(ln, NS.a, "noFill")) return "none";
  const solid = directChild(ln, NS.a, "solidFill");
  const color = solid ? resolveDrawingMlColor(parseColorChoice(solid), ctx.theme, ctx.clrMap) : null;
  if (!color) return null;
  return `1px solid ${color}`;
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
