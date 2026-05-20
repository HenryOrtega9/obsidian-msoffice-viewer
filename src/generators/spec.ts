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
    const level = obj.level;
    if (level !== 1 && level !== 2 && level !== 3)
      throw new SpecValidationError(`${path}.level: must be 1, 2, or 3`);
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
    return { type: t, items: obj.items };
  }
  if (t === "table") {
    const rows = obj.rows;
    if (!Array.isArray(rows) || !rows.every((r) => isStrArr(r)))
      throw new SpecValidationError(`${path}.rows: must be a 2D array of strings`);
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
  // Strip ```json ... ``` or ``` ... ``` fences if present
  const fence = s.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  if (fence) s = fence[1].trim();
  // If extra prose surrounds the JSON, try to slice from first { to last }
  if (!s.startsWith("{")) {
    const first = s.indexOf("{");
    const last = s.lastIndexOf("}");
    if (first !== -1 && last > first) s = s.slice(first, last + 1);
  }
  return JSON.parse(s);
}
