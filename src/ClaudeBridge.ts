/*
 * ClaudeBridge - integration point with the `claude-cli-chat` Obsidian plugin.
 *
 * IMPORTANT: confirm the API surface of claude-cli-chat by reading
 * <vault>/.obsidian/plugins/claude-cli-chat/main.js at install time.
 * The probing order (instance method then command id) covers the two shapes
 * most likely to exist. If neither is present, surface a clear error.
 */

import { App, Notice } from "obsidian";
import type {
  ClaudeEditResponse,
  ClaudeEditType,
  FidelityHint,
  StructuralLocator,
} from "./types";

type ClaudeCliPluginInstance = {
  runHeadlessPrompt?: (
    systemPrompt: string,
    userPrompt: string,
  ) => Promise<string>;
  sendPrompt?: (prompt: string, opts?: unknown) => Promise<string>;
};

const SYSTEM_PROMPT = `You are an editor helping rewrite a span of text inside a Microsoft Word (.docx) document opened in Obsidian.

Respond with a single JSON object. No prose, no markdown fences. Schema:
{
  "type": "replace_text" | "rewrite_section" | "insert_after" | "delete",
  "target_locator": <copy the locator object provided verbatim>,
  "new_content": "<the rewritten text>",
  "fidelity_hint": "text_only" | "structural"
}

Use "replace_text" with "text_only" when the change is a pure rewrite of the selected text and does not introduce new paragraphs, lists, headings, tables, or formatting changes.
Use "rewrite_section" or "insert_after" with "structural" when the edit adds/removes paragraphs, headings, lists, or otherwise changes document structure.
Use "delete" with "new_content": "" when the user wants the span removed.`;

function buildUserPrompt(
  locator: StructuralLocator,
  instruction: string,
): string {
  return [
    `Instruction: ${instruction}`,
    "",
    "Selected text:",
    locator.selectedText,
    "",
    "Containing paragraph:",
    locator.paragraphText,
    "",
    "Surrounding context (prior/next paragraph snippets):",
    locator.surroundingContext,
    "",
    "Locator (echo this back verbatim in target_locator):",
    JSON.stringify(locator),
    "",
    `Note: the selection ${
      locator.crossesFormatting
        ? "spans multiple formatting runs - prefer structural unless the change is a trivial rewrite"
        : "is within a single formatting run - text_only is usually safe"
    }.`,
  ].join("\n");
}

const VALID_TYPES: ReadonlySet<ClaudeEditType> = new Set([
  "replace_text",
  "rewrite_section",
  "insert_after",
  "delete",
]);

const VALID_HINTS: ReadonlySet<FidelityHint> = new Set([
  "text_only",
  "structural",
]);

function tryExtractJson(raw: string): string | null {
  const trimmed = raw.trim();
  if (trimmed.startsWith("{")) return trimmed;
  const fenceMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fenceMatch) return fenceMatch[1].trim();
  const braceStart = trimmed.indexOf("{");
  const braceEnd = trimmed.lastIndexOf("}");
  if (braceStart >= 0 && braceEnd > braceStart) {
    return trimmed.slice(braceStart, braceEnd + 1);
  }
  return null;
}

function validateResponse(
  obj: unknown,
  fallbackLocator: StructuralLocator,
): ClaudeEditResponse | null {
  if (!obj || typeof obj !== "object") return null;
  const r = obj as Record<string, unknown>;
  if (typeof r.type !== "string" || !VALID_TYPES.has(r.type as ClaudeEditType)) {
    return null;
  }
  if (
    typeof r.fidelity_hint !== "string" ||
    !VALID_HINTS.has(r.fidelity_hint as FidelityHint)
  ) {
    return null;
  }
  if (typeof r.new_content !== "string") return null;
  const target =
    r.target_locator && typeof r.target_locator === "object"
      ? (r.target_locator as StructuralLocator)
      : fallbackLocator;
  return {
    type: r.type as ClaudeEditType,
    fidelity_hint: r.fidelity_hint as FidelityHint,
    new_content: r.new_content,
    target_locator: target,
  };
}

export class ClaudeBridge {
  constructor(private app: App) {}

  private getPlugin(): ClaudeCliPluginInstance | null {
    const plugins = (this.app as unknown as {
      plugins?: {
        plugins?: Record<string, ClaudeCliPluginInstance | undefined>;
        enabledPlugins?: Set<string>;
      };
    }).plugins;
    if (!plugins?.plugins) return null;
    return plugins.plugins["claude-cli-chat"] ?? null;
  }

  isAvailable(): boolean {
    const p = this.getPlugin();
    if (!p) return false;
    return (
      typeof p.runHeadlessPrompt === "function" ||
      typeof p.sendPrompt === "function"
    );
  }

  private async callClaude(
    systemPrompt: string,
    userPrompt: string,
  ): Promise<string> {
    const p = this.getPlugin();
    if (!p) {
      throw new Error(
        "claude-cli-chat plugin is not installed or enabled. Enable it in Settings -> Community plugins.",
      );
    }
    if (typeof p.runHeadlessPrompt === "function") {
      return await p.runHeadlessPrompt(systemPrompt, userPrompt);
    }
    if (typeof p.sendPrompt === "function") {
      const combined = `${systemPrompt}\n\n---\n\n${userPrompt}`;
      return await p.sendPrompt(combined);
    }
    throw new Error(
      "claude-cli-chat does not expose a programmatic prompt method. Upgrade it to a build that exports runHeadlessPrompt.",
    );
  }

  async askClaude(
    locator: StructuralLocator,
    instruction: string,
  ): Promise<ClaudeEditResponse> {
    const userPrompt = buildUserPrompt(locator, instruction);
    let raw: string;
    try {
      raw = await this.callClaude(SYSTEM_PROMPT, userPrompt);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      new Notice(`Claude bridge error: ${message}`);
      throw err;
    }

    const jsonText = tryExtractJson(raw);
    if (jsonText) {
      try {
        const parsed = JSON.parse(jsonText);
        const validated = validateResponse(parsed, locator);
        if (validated) return validated;
      } catch {
        // fall through to synthetic
      }
    }

    return {
      type: "rewrite_section",
      fidelity_hint: "structural",
      new_content: raw.trim(),
      target_locator: locator,
    };
  }
}
