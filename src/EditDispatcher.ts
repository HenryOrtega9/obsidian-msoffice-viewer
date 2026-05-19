import { App, Notice, TFile } from "obsidian";
import { ooxmlTextReplace } from "./edits/ooxmlTextReplace";
import { markdownRoundtrip } from "./edits/markdownRoundtrip";
import { BackupManager } from "./BackupManager";
import {
  ClaudeEditResponse,
  EditOutcome,
  EscalateToRoundtrip,
  StructuralLocator,
} from "./types";
import type { DocxClaudeSettings } from "./settings";

export interface DispatcherContext {
  app: App;
  pluginDir: string;
  backup: BackupManager;
  settings: DocxClaudeSettings;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const ab = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(ab).set(bytes);
  return ab;
}

function canUseOoxmlPath(
  response: ClaudeEditResponse,
  locator: StructuralLocator,
): boolean {
  return (
    response.type === "replace_text" &&
    response.fidelity_hint === "text_only" &&
    !locator.crossesFormatting
  );
}

export async function dispatch(
  ctx: DispatcherContext,
  file: TFile,
  response: ClaudeEditResponse,
  locator: StructuralLocator,
): Promise<EditOutcome> {
  await ctx.backup.snapshot(file);

  let pathUsed: "ooxml" | "roundtrip" = "roundtrip";
  let warning: string | undefined;
  let bytes: Uint8Array;

  if (canUseOoxmlPath(response, locator)) {
    try {
      const buf = await ctx.app.vault.readBinary(file);
      bytes = await ooxmlTextReplace(buf, locator, response.new_content);
      pathUsed = "ooxml";
    } catch (err) {
      if (err instanceof EscalateToRoundtrip) {
        new Notice(`Falling back to round-trip: ${err.message}`);
        bytes = await runRoundtrip(ctx, file, locator, response);
      } else {
        throw err;
      }
    }
  } else {
    bytes = await runRoundtrip(ctx, file, locator, response);
  }

  await ctx.app.vault.modifyBinary(file, toArrayBuffer(bytes));

  if (pathUsed === "roundtrip") {
    const count = ctx.backup.recordRoundtrip(file.path);
    if (count >= ctx.settings.warnAfterNRoundtrips) {
      warning = `This file has been round-tripped ${count} times this session. Fidelity may have drifted.`;
    }
  }

  return { pathUsed, bytesWritten: bytes.byteLength, warning };
}

async function runRoundtrip(
  ctx: DispatcherContext,
  file: TFile,
  locator: StructuralLocator,
  response: ClaudeEditResponse,
): Promise<Uint8Array> {
  return markdownRoundtrip(
    {
      app: ctx.app,
      file,
      pluginDir: ctx.pluginDir,
      pandocPath: ctx.settings.pandocPath,
    },
    locator,
    response,
  );
}
