import { App, TFile, normalizePath } from "obsidian";
import * as mammoth from "mammoth";
import { spawn } from "child_process";
import { promises as fs } from "fs";
import * as path from "path";
import * as os from "os";
import type { ClaudeEditResponse, StructuralLocator } from "../types";

export interface RoundtripContext {
  app: App;
  file: TFile;
  pluginDir: string;
  pandocPath: string;
}

async function ensureTmpDir(app: App, pluginDir: string): Promise<string> {
  const rel = normalizePath(`${pluginDir}/tmp`);
  const exists = await app.vault.adapter.exists(rel);
  if (!exists) await app.vault.adapter.mkdir(rel);
  const adapter = app.vault.adapter as unknown as { basePath?: string };
  if (adapter.basePath) {
    const full = path.join(adapter.basePath, rel);
    await fs.mkdir(full, { recursive: true });
    return full;
  }
  const fallback = path.join(os.tmpdir(), "obsidian-docx-claude");
  await fs.mkdir(fallback, { recursive: true });
  return fallback;
}

function spliceParagraph(
  markdown: string,
  oldParagraphText: string,
  newContent: string,
  response: ClaudeEditResponse,
): { result: string; matched: boolean } {
  const needle = oldParagraphText.trim();
  if (!needle) {
    return { result: `${markdown}\n\n${newContent}`, matched: false };
  }
  const idx = markdown.indexOf(needle);
  if (idx === -1) {
    return {
      result: `${markdown}\n\n<!-- claude-edit (unmatched splice) -->\n${newContent}\n`,
      matched: false,
    };
  }
  const endOfPara = markdown.indexOf("\n\n", idx + needle.length);
  const tail = endOfPara === -1 ? markdown.length : endOfPara;
  switch (response.type) {
    case "delete":
      return {
        result: markdown.slice(0, idx) + markdown.slice(tail).replace(/^\n+/, ""),
        matched: true,
      };
    case "insert_after":
      return {
        result:
          markdown.slice(0, tail) + `\n\n${newContent}` + markdown.slice(tail),
        matched: true,
      };
    case "replace_text":
    case "rewrite_section":
    default:
      return {
        result: markdown.slice(0, idx) + newContent + markdown.slice(tail),
        matched: true,
      };
  }
}

function runPandoc(
  pandocPath: string,
  args: string[],
  cwd: string,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(pandocPath, args, { cwd });
    let stderr = "";
    child.stderr.on("data", (d) => (stderr += d.toString()));
    child.on("error", (err) =>
      reject(new Error(`Failed to spawn pandoc: ${err.message}`)),
    );
    child.on("close", (code) => {
      if (code === 0) resolve();
      else
        reject(
          new Error(
            `pandoc exited with code ${code}${stderr ? `: ${stderr}` : ""}`,
          ),
        );
    });
  });
}

export async function markdownRoundtrip(
  ctx: RoundtripContext,
  locator: StructuralLocator,
  response: ClaudeEditResponse,
): Promise<Uint8Array> {
  const originalBuffer = await ctx.app.vault.readBinary(ctx.file);
  const tmpDir = await ensureTmpDir(ctx.app, ctx.pluginDir);

  const refDocPath = path.join(tmpDir, `ref-${Date.now()}.docx`);
  await fs.writeFile(refDocPath, Buffer.from(originalBuffer));

  const md = await mammoth.convertToMarkdown({
    buffer: Buffer.from(originalBuffer),
  });
  const markdown = md.value;

  const { result: newMarkdown } = spliceParagraph(
    markdown,
    locator.paragraphText,
    response.new_content,
    response,
  );

  const inPath = path.join(tmpDir, `in-${Date.now()}.md`);
  const outPath = path.join(tmpDir, `out-${Date.now()}.docx`);
  await fs.writeFile(inPath, newMarkdown, "utf8");

  try {
    await runPandoc(
      ctx.pandocPath,
      [
        "-f",
        "markdown",
        "-t",
        "docx",
        "--reference-doc",
        refDocPath,
        "-o",
        outPath,
        inPath,
      ],
      tmpDir,
    );
    const bytes = await fs.readFile(outPath);
    return new Uint8Array(bytes);
  } finally {
    await Promise.allSettled([
      fs.unlink(inPath),
      fs.unlink(outPath),
      fs.unlink(refDocPath),
    ]);
  }
}
