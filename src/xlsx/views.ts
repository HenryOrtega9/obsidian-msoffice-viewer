import JSZip from "jszip";

// LibreOffice's SinglePageSheets PDF export (one un-paginated page per sheet —
// the high-fidelity xlsx render path) takes its PAINT ORIGIN from the
// worksheet's saved scroll position instead of the sheet's top-left. A workbook
// saved scrolled down — sheetView/@topLeftCell="A39", or a frozen/split pane
// whose own pane/@topLeftCell carries the offset — then renders a page sized to
// the full used range but with every row above the saved offset left BLANK:
// their height is reserved, they're just never painted (the cause of the
// "blank top band, correct bottom" xlsx preview). This is an unfixed upstream
// defect (tdf#164683 / tdf#155351, both still NEW as of mid-2026), so there is
// no LibreOffice version to upgrade into and no calc_pdf_Export FilterData flag
// that resets the viewport while SinglePageSheets is on. We neutralize it at the
// source: reset the saved scroll offset to A1 in every worksheet before staging
// the bytes for soffice. The frozen pane element itself (xSplit/ySplit/state/
// activePane) is left intact so the frozen header band still renders — only the
// scroll offset moves (Variant A: zeroing topLeftCell alone fixes it; Variant B:
// removing the pane alone does NOT). Everything else (showGridLines, view,
// zoomScale, selection, drawings/charts, and all non-worksheet parts) is left
// byte-untouched to preserve fidelity.
//
// Returns the rewritten workbook bytes, or null when no worksheet carries a
// scroll offset or anything goes wrong — so the caller converts its input bytes
// unchanged. Normalization must never break a conversion. This is applied only
// to the staged copy fed to LibreOffice, never to the user's stored file.
export async function normalizeSheetViews(
  bytes: Uint8Array,
): Promise<Uint8Array | null> {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const sheetNames = Object.keys(zip.files).filter((n) =>
      /^xl\/worksheets\/sheet[^/]*\.xml$/.test(n),
    );
    if (sheetNames.length === 0) return null;

    let changed = false;
    for (const n of sheetNames) {
      if (await editEntry(zip, n, resetScrollOffset)) changed = true;
    }
    // Nothing was scrolled — return null so the caller keeps its input bytes
    // and we skip a pointless re-zip.
    if (!changed) return null;

    return await zip.generateAsync({ type: "uint8array" });
  } catch {
    return null;
  }
}

// Reset the saved scroll offset (sheetView/@topLeftCell and any child
// pane/@topLeftCell) to "A1" within each active <sheetView> element of one
// worksheet's XML. Scoped to the <sheetView> element (start tag through its
// </sheetView>, or the self-closing form) rather than to a global <pane> match,
// so panes saved inside <customSheetViews> — which never drive the export — stay
// byte-untouched. Within a sheetView, topLeftCell only ever appears on the
// sheetView start tag and its child <pane>, so zeroing every topLeftCell in the
// matched block is exact. The rewrite is global (split-window sheets can hold
// several <sheetView> elements), handles both quote styles, and only rewrites an
// EXISTING attribute — a sheet defaulting to A1 has no attribute and is correctly
// a no-op, so no malformed XML is produced. Namespace-agnostic: topLeftCell
// carries no prefix in the worksheet namespace.
function resetScrollOffset(xml: string): string {
  // Cheap bail: nothing to do when no scroll offset is stored anywhere.
  // Whitespace-tolerant (matches the \s*=\s* the replacement regexes accept) so
  // a non-Excel producer emitting `topLeftCell =` can't slip past the fast-path.
  if (!/topLeftCell\s*=/.test(xml)) return xml;
  const zeroTopLeft = (block: string): string =>
    block
      .replace(/\btopLeftCell\s*=\s*"[^"]*"/g, 'topLeftCell="A1"')
      .replace(/\btopLeftCell\s*=\s*'[^']*'/g, 'topLeftCell="A1"');
  return xml.replace(
    /<(?:[\w.-]+:)?sheetView\b[^>]*\/>|<(?:[\w.-]+:)?sheetView\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?sheetView>/g,
    zeroTopLeft,
  );
}

// Returns true when the transform actually changed the entry (so the caller can
// tell whether the workbook needs re-zipping). Mirrors flatten.ts's editEntry.
async function editEntry(
  zip: JSZip,
  name: string,
  transform: (xml: string) => string,
): Promise<boolean> {
  const file = zip.file(name);
  if (!file) return false;
  const xml = await file.async("string");
  const next = transform(xml);
  if (next === xml) return false;
  zip.file(name, next);
  return true;
}
