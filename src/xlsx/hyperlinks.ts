import type ExcelJS from "exceljs";

export interface SheetHyperlink {
  hyperlink: string;
  tooltip?: string;
}

export interface InternalLinkTarget {
  sheet: string;
  address: string;
}

export type InternalLinkHandler = (target: InternalLinkTarget) => void;

// Build a quick-lookup map for sheet-level hyperlinks. ExcelJS stores some
// hyperlinks on cell.value, others in ws.model.hyperlinks — check both.
export function collectSheetHyperlinks(ws: ExcelJS.Worksheet): Map<string, SheetHyperlink> {
  const out = new Map<string, SheetHyperlink>();
  const list = (ws as unknown as {
    model?: { hyperlinks?: Array<{ address?: string; hyperlink?: string; tooltip?: string }> };
  }).model?.hyperlinks;
  if (!Array.isArray(list)) return out;
  for (const h of list) {
    if (h.address && h.hyperlink) {
      out.set(h.address.toUpperCase(), { hyperlink: h.hyperlink, tooltip: h.tooltip });
    }
  }
  return out;
}

// Resolve a cell's hyperlink target by checking cell.value shape first, then
// the sheet-level map keyed by cell address (e.g. "B5").
export function getCellHyperlink(
  cell: ExcelJS.Cell,
  sheetMap: Map<string, SheetHyperlink>,
): SheetHyperlink | null {
  const v = cell.value as unknown;
  if (v && typeof v === "object" && "hyperlink" in v) {
    const obj = v as { hyperlink?: unknown; tooltip?: unknown };
    if (typeof obj.hyperlink === "string") {
      return {
        hyperlink: obj.hyperlink,
        tooltip: typeof obj.tooltip === "string" ? obj.tooltip : undefined,
      };
    }
  }
  const fromMap = sheetMap.get(cell.address?.toUpperCase() ?? "");
  return fromMap ?? null;
}

// Replace td.children with an <a> wrapping the existing content. Preserves any
// pre-rendered spans (rich text); pass `inheritStyling: true` so the link
// inherits the cell font color rather than the browser default blue.
export function wrapInHyperlink(
  td: HTMLTableCellElement,
  link: SheetHyperlink,
  opts: { onInternal?: InternalLinkHandler } = {},
): void {
  const url = link.hyperlink;
  const internal = parseInternalLink(url);

  // Move existing children into the anchor.
  const a = document.createElement("a");
  a.className = "docx-claude-xlsx-link";
  if (link.tooltip) a.title = link.tooltip;

  if (internal) {
    a.href = "#";
    a.addEventListener("click", (ev) => {
      ev.preventDefault();
      opts.onInternal?.(internal);
    });
  } else {
    a.href = url;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
  }

  while (td.firstChild) a.appendChild(td.firstChild);
  // If the cell was empty, show the URL as fallback text.
  if (!a.firstChild) a.setText(internal ? `${internal.sheet}!${internal.address}` : url);
  td.appendChild(a);
}

// Recognize forms like `#'Sheet Name'!A1`, `#Sheet1!B5`, absolute refs like
// `#Sheet1!$A$1`, ranges like `#Sheet1!$A$1:$B$10` (reduced to the top-left
// cell), or sheet-relative targets ExcelJS sometimes hands us without the
// leading hash. The address part allows optional $ on row/column.
function parseInternalLink(url: string): InternalLinkTarget | null {
  const stripped = url.startsWith("#") ? url.slice(1) : url;
  // Strip absolute-ref $ and reduce a range to its top-left cell so
  // followInternalLink's colNum/cellMatch keep working unchanged.
  const clean = (a: string): string => a.replace(/\$/g, "").split(":")[0];
  const addr = "\\$?[A-Z]+\\$?\\d+(?::\\$?[A-Z]+\\$?\\d+)?";
  // Quoted sheet name: '...' allowing escaped quote pair ''
  const quoted = stripped.match(new RegExp(`^'((?:[^']|'')+)'!(${addr})$`));
  if (quoted) return { sheet: quoted[1].replace(/''/g, "'"), address: clean(quoted[2]) };
  // Unquoted sheet name (no spaces or special chars).
  const plain = stripped.match(new RegExp(`^([A-Za-z_][A-Za-z0-9_]*)!(${addr})$`));
  if (plain) return { sheet: plain[1], address: clean(plain[2]) };
  return null;
}
