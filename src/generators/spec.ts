export type FileKind = "docx" | "pptx" | "xlsx";

export type DocxBlock =
  | { type: "heading"; level: 1 | 2 | 3; text: string }
  | { type: "paragraph"; text: string }
  | { type: "bullets"; items: string[] }
  | { type: "numbered"; items: string[] }
  | { type: "table"; rows: string[][] };

export interface DocxSpec {
  kind: "docx";
  title?: string;
  blocks: DocxBlock[];
}

export interface PptxSlide {
  title?: string;
  bullets?: string[];
  body?: string;
  notes?: string;
}

export interface PptxSpec {
  kind: "pptx";
  title?: string;
  slides: PptxSlide[];
}

export type XlsxCell = string | number | boolean | null;

export interface XlsxSheet {
  name: string;
  rows: XlsxCell[][];
}

export interface XlsxSpec {
  kind: "xlsx";
  sheets: XlsxSheet[];
}

export type CreateSpec = DocxSpec | PptxSpec | XlsxSpec;

export class SpecValidationError extends Error {}

function isStr(v: unknown): v is string {
  return typeof v === "string";
}

function isStrArr(v: unknown): v is string[] {
  return Array.isArray(v) && v.every(isStr);
}

function validateDocxBlock(b: unknown, path: string): DocxBlock {
  if (!b || typeof b !== "object") throw new SpecValidationError(`${path}: expected object`);
  const obj = b as Record<string, unknown>;
  const t = obj.type;
  if (t === "heading") {
    // Accept numeric strings ("1") and clamp out-of-range levels (1..3) so a
    // pedantic JSON shape doesn't sink the whole document.
    const rawLevel = obj.level;
    const n =
      typeof rawLevel === "number"
        ? rawLevel
        : typeof rawLevel === "string"
          ? Number(rawLevel)
          : NaN;
    if (!Number.isFinite(n)) {
      throw new SpecValidationError(`${path}.level: must be a number`);
    }
    const level = (Math.max(1, Math.min(3, Math.round(n))) as 1 | 2 | 3);
    if (!isStr(obj.text)) throw new SpecValidationError(`${path}.text: must be a string`);
    return { type: "heading", level, text: obj.text };
  }
  if (t === "paragraph") {
    if (!isStr(obj.text)) throw new SpecValidationError(`${path}.text: must be a string`);
    return { type: "paragraph", text: obj.text };
  }
  if (t === "bullets" || t === "numbered") {
    if (!isStrArr(obj.items))
      throw new SpecValidationError(`${path}.items: must be an array of strings`);
    if (obj.items.length === 0)
      throw new SpecValidationError(`${path}.items: must not be empty`);
    return { type: t, items: obj.items };
  }
  if (t === "table") {
    const rows = obj.rows;
    if (!Array.isArray(rows) || !rows.every((r) => isStrArr(r)))
      throw new SpecValidationError(`${path}.rows: must be a 2D array of strings`);
    if (rows.length === 0)
      throw new SpecValidationError(`${path}.rows: must not be empty`);
    const width = (rows[0] as string[]).length;
    if (width === 0)
      throw new SpecValidationError(`${path}.rows[0]: header row must not be empty`);
    if (!(rows as string[][]).every((r) => r.length === width)) {
      throw new SpecValidationError(
        `${path}.rows: every row must have ${width} cells (header width)`,
      );
    }
    return { type: "table", rows: rows as string[][] };
  }
  throw new SpecValidationError(`${path}.type: unknown block type "${String(t)}"`);
}

export function validateSpec(kind: FileKind, raw: unknown): CreateSpec {
  if (!raw || typeof raw !== "object") throw new SpecValidationError("spec: expected object");
  const obj = raw as Record<string, unknown>;

  if (kind === "docx") {
    const blocks = obj.blocks;
    if (!Array.isArray(blocks))
      throw new SpecValidationError("spec.blocks: must be an array");
    const validated = blocks.map((b, i) => validateDocxBlock(b, `spec.blocks[${i}]`));
    return {
      kind: "docx",
      title: isStr(obj.title) ? obj.title : undefined,
      blocks: validated,
    };
  }

  if (kind === "pptx") {
    const slides = obj.slides;
    if (!Array.isArray(slides) || slides.length === 0)
      throw new SpecValidationError("spec.slides: must be a non-empty array");
    const validated: PptxSlide[] = slides.map((s, i) => {
      if (!s || typeof s !== "object")
        throw new SpecValidationError(`spec.slides[${i}]: expected object`);
      const sl = s as Record<string, unknown>;
      const out: PptxSlide = {};
      if (isStr(sl.title)) out.title = sl.title;
      if (Array.isArray(sl.bullets)) {
        if (!isStrArr(sl.bullets))
          throw new SpecValidationError(`spec.slides[${i}].bullets: must be string[]`);
        out.bullets = sl.bullets;
      }
      if (isStr(sl.body)) out.body = sl.body;
      if (isStr(sl.notes)) out.notes = sl.notes;
      return out;
    });
    return {
      kind: "pptx",
      title: isStr(obj.title) ? obj.title : undefined,
      slides: validated,
    };
  }

  // xlsx
  const sheets = obj.sheets;
  if (!Array.isArray(sheets) || sheets.length === 0)
    throw new SpecValidationError("spec.sheets: must be a non-empty array");
  const validated: XlsxSheet[] = sheets.map((s, i) => {
    if (!s || typeof s !== "object")
      throw new SpecValidationError(`spec.sheets[${i}]: expected object`);
    const sh = s as Record<string, unknown>;
    if (!isStr(sh.name))
      throw new SpecValidationError(`spec.sheets[${i}].name: must be a string`);
    const rows = sh.rows;
    if (!Array.isArray(rows))
      throw new SpecValidationError(`spec.sheets[${i}].rows: must be a 2D array`);
    const validatedRows: XlsxCell[][] = rows.map((r, j) => {
      if (!Array.isArray(r))
        throw new SpecValidationError(`spec.sheets[${i}].rows[${j}]: must be an array`);
      return r.map((cell, k) => {
        if (cell === null) return null;
        const t = typeof cell;
        if (t === "string" || t === "number" || t === "boolean") return cell as XlsxCell;
        throw new SpecValidationError(
          `spec.sheets[${i}].rows[${j}][${k}]: must be string|number|boolean|null`,
        );
      });
    });
    return { name: sh.name, rows: validatedRows };
  });
  return { kind: "xlsx", sheets: validated };
}

export function extractJsonFromResponse(raw: string): unknown {
  let s = raw.trim();
  // Strip a ```json … ``` / ``` … ``` fence wherever it appears, even if
  // prose follows the fence. Non-anchored so trailing model commentary
  // doesn't defeat the match.
  const fence = s.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (fence) s = fence[1].trim();
  // If we still don't start with `{`, walk the string and find the first
  // top-level balanced `{ … }` (respecting strings + escapes). This is more
  // robust than slice(first-`{`, last-`}`) when there are stray braces in
  // surrounding prose or trailing remarks.
  if (!s.startsWith("{")) {
    const sliced = extractFirstJsonObject(s);
    if (sliced !== null) s = sliced;
  }
  return JSON.parse(s);
}

function extractFirstJsonObject(s: string): string | null {
  const start = s.indexOf("{");
  if (start === -1) return null;
  let depth = 0;
  let inString = false;
  let escape = false;
  for (let i = start; i < s.length; i++) {
    const ch = s[i];
    if (inString) {
      if (escape) escape = false;
      else if (ch === "\\") escape = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return s.slice(start, i + 1);
    }
  }
  return null;
}
