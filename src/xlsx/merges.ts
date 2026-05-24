import type ExcelJS from "exceljs";

export interface MergeRect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export function collectMerges(ws: ExcelJS.Worksheet): MergeRect[] {
  const merges: MergeRect[] = [];
  const list = (ws as unknown as { model?: { merges?: string[] } }).model?.merges;
  if (!Array.isArray(list)) return merges;
  for (const addr of list) {
    const m = parseMergeRange(addr);
    if (m) merges.push(m);
  }
  return merges;
}

export function parseMergeRange(addr: string): MergeRect | null {
  const m = addr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return null;
  return {
    top: parseInt(m[2], 10),
    left: colNum(m[1]),
    bottom: parseInt(m[4], 10),
    right: colNum(m[3]),
  };
}

export function computeMergeSkipMap(merges: MergeRect[]): Set<string> {
  const skip = new Set<string>();
  for (const m of merges) {
    for (let r = m.top; r <= m.bottom; r++) {
      for (let c = m.left; c <= m.right; c++) {
        if (r === m.top && c === m.left) continue;
        skip.add(`${r}:${c}`);
      }
    }
  }
  return skip;
}

export function colLetter(c: number): string {
  let s = "";
  while (c > 0) {
    const rem = (c - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colNum(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
