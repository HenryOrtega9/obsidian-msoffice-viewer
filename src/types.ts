export interface StructuralLocator {
  paragraphIndex: number;
  startOffset: number;
  endOffset: number;
  selectedText: string;
  paragraphText: string;
  surroundingContext: string;
  crossesFormatting: boolean;
}

export type ClaudeEditType =
  | "replace_text"
  | "rewrite_section"
  | "insert_after"
  | "delete";

export type FidelityHint = "text_only" | "structural";

export interface ClaudeEditResponse {
  type: ClaudeEditType;
  target_locator: StructuralLocator;
  new_content: string;
  fidelity_hint: FidelityHint;
}

export interface EditOutcome {
  pathUsed: "ooxml" | "roundtrip";
  bytesWritten: number;
  warning?: string;
}

export class EscalateToRoundtrip extends Error {
  constructor(reason: string) {
    super(reason);
    this.name = "EscalateToRoundtrip";
  }
}
