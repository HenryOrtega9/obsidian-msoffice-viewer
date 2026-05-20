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
    let name = sanitizeSheetName(sheet.name, `Sheet${i + 1}`);
    let suffix = 2;
    while (usedNames.has(name)) {
      const base = sanitizeSheetName(sheet.name, `Sheet${i + 1}`);
      const candidate = `${base} (${suffix++})`;
      name = candidate.length > 31 ? candidate.slice(0, 31) : candidate;
    }
    usedNames.add(name);

    const ws = wb.addWorksheet(name);
    sheet.rows.forEach((row, rowIdx) => {
      const added = ws.addRow(row.map((c) => (c === null ? null : c)));
      if (rowIdx === 0) added.font = { bold: true };
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
