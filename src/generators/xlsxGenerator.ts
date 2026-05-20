import ExcelJS from "exceljs";
import { XlsxSpec } from "./spec";

const ILLEGAL_SHEET_CHARS = /[:\\/?*\[\]]/g;

function sanitizeSheetName(name: string, fallback: string): string {
  let s = name.replace(ILLEGAL_SHEET_CHARS, " ").trim();
  if (!s) s = fallback;
  if (s.length > 31) s = s.slice(0, 31);
  return s;
}

export async function buildXlsx(spec: XlsxSpec): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const usedNames = new Set<string>();

  spec.sheets.forEach((sheet, i) => {
    const baseSanitized = sanitizeSheetName(sheet.name, `Sheet${i + 1}`);
    let name = baseSanitized;
    let suffix = 2;
    while (usedNames.has(name)) {
      const tag = ` (${suffix++})`;
      // Reserve room for the suffix before slicing so the unique part can't
      // be chopped off (would otherwise infinite-loop on 31-char duplicates).
      const stem = baseSanitized.slice(0, Math.max(1, 31 - tag.length));
      name = stem + tag;
    }
    usedNames.add(name);

    const ws = wb.addWorksheet(name);
    sheet.rows.forEach((row, rowIdx) => {
      const added = ws.addRow(row.map((c) => (c === null ? null : c)));
      if (rowIdx === 0) {
        // Row.font setter doesn't apply to cells whose value is null. Iterate
        // explicitly so a header row like ["A", null, "B"] still bolds A and B.
        added.eachCell({ includeEmpty: true }, (cell) => {
          cell.font = { bold: true };
        });
      }
    });

    const colCount = sheet.rows[0]?.length ?? 0;
    for (let c = 1; c <= colCount; c++) {
      const col = ws.getColumn(c);
      let max = 10;
      col.eachCell({ includeEmpty: false }, (cell) => {
        const len = String(cell.value ?? "").length;
        if (len > max) max = len;
      });
      col.width = Math.min(max + 2, 60);
    }
  });

  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer | Uint8Array;
  if (buf instanceof ArrayBuffer) return buf;
  // ExcelJS returns a Node Buffer in node environments; copy into a plain ArrayBuffer.
  const u8 = buf as Uint8Array;
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
}
