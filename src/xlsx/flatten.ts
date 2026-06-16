import JSZip from "jszip";

// LibreOffice rebuilds Excel PivotTables through its own DataPilot engine on
// import. That rebuild drops the pivot field's stored number formats (so
// "$1,325.52" / "1.77" export as raw floats like 1325.52094736842) and renames
// "Grand Total" to "Total Result". The worksheet XML already carries the fully
// cached pivot result cells WITH their stored number formats and indentation,
// so deleting the pivotTable / pivotCache parts makes LibreOffice render those
// static cells faithfully instead of regenerating a DataPilot.
//
// Returns the flattened workbook bytes, or null when the workbook has no pivot
// tables or anything goes wrong — so the caller converts the original bytes
// unchanged. Flattening must never break a conversion.
export async function flattenPivotTables(
  bytes: Uint8Array,
): Promise<Uint8Array | null> {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const names = Object.keys(zip.files);
    if (!names.some((n) => /^xl\/pivotTables\/.+\.xml$/.test(n))) return null;

    // 1) Drop the pivotTable and pivotCache parts entirely.
    for (const n of names) {
      if (/^xl\/pivotTables\//.test(n) || /^xl\/pivotCache\//.test(n)) {
        zip.remove(n);
      }
    }

    // 2) Strip the pivotTable relationship from every worksheet's .rels so each
    //    sheet renders its cached static cells rather than a rebuilt DataPilot.
    for (const n of names) {
      if (/^xl\/worksheets\/_rels\/.+\.rels$/.test(n)) {
        await editEntry(zip, n, (xml) =>
          xml.replace(/<Relationship\b[^>]*\bpivotTable\b[^>]*\/>/g, ""),
        );
      }
    }

    // 3) Remove the workbook-level <pivotCaches> element, its relationships, and
    //    the [Content_Types] overrides so no dangling references remain (a
    //    dangling reference can make LibreOffice reject the file).
    await editEntry(zip, "xl/workbook.xml", (xml) =>
      xml.replace(/<pivotCaches>[\s\S]*?<\/pivotCaches>/g, ""),
    );
    await editEntry(zip, "xl/_rels/workbook.xml.rels", (xml) =>
      xml.replace(/<Relationship\b[^>]*pivotCache[^>]*\/>/g, ""),
    );
    await editEntry(zip, "[Content_Types].xml", (xml) =>
      xml.replace(/<Override\b[^>]*pivot(?:Table|Cache)[^>]*\/>/g, ""),
    );

    return await zip.generateAsync({ type: "uint8array" });
  } catch {
    return null;
  }
}

async function editEntry(
  zip: JSZip,
  name: string,
  transform: (xml: string) => string,
): Promise<void> {
  const file = zip.file(name);
  if (!file) return;
  const xml = await file.async("string");
  const next = transform(xml);
  if (next !== xml) zip.file(name, next);
}
