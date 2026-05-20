import { CreateSpec } from "./spec";
import { buildDocx } from "./docxGenerator";
import { buildPptx } from "./pptxGenerator";
import { buildXlsx } from "./xlsxGenerator";

export async function buildFile(spec: CreateSpec): Promise<ArrayBuffer> {
  switch (spec.kind) {
    case "docx":
      return buildDocx(spec);
    case "pptx":
      return buildPptx(spec);
    case "xlsx":
      return buildXlsx(spec);
  }
}

export * from "./spec";
export { systemPromptFor } from "./prompts";
