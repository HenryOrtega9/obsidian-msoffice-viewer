import * as pdfjsLib from "pdfjs-dist";
import { App, FileSystemAdapter } from "obsidian";
import { promises as fs } from "node:fs";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { createHash } from "node:crypto";
import * as path from "node:path";
import * as os from "node:os";

declare const __PDFJS_WORKER_SOURCE__: string;

// pdfjs-dist v5.7 calls JS methods that ship in very recent V8 builds (Map
// getOrInsertComputed, Math.sumPrecise, Set.prototype.intersection). Obsidian's
// bundled Electron is older. We install the polyfills both on the main thread
// (via new Function on the source below) and prepend the same source to the
// worker file we write to disk so the worker context also has them.
const RUNTIME_POLYFILLS = `
if (typeof Map.prototype.getOrInsertComputed !== "function") {
  Object.defineProperty(Map.prototype, "getOrInsertComputed", {
    value: function (key, computeFn) {
      if (this.has(key)) return this.get(key);
      const v = computeFn(key);
      this.set(key, v);
      return v;
    },
    writable: true,
    configurable: true,
  });
}
if (typeof Math.sumPrecise !== "function") {
  Object.defineProperty(Math, "sumPrecise", {
    value: function (iter) {
      let sum = 0;
      let comp = 0;
      for (const n of iter) {
        const y = n - comp;
        const t = sum + y;
        comp = (t - sum) - y;
        sum = t;
      }
      return sum;
    },
    writable: true,
    configurable: true,
  });
}
if (typeof Set.prototype.intersection !== "function") {
  Object.defineProperty(Set.prototype, "intersection", {
    value: function (other) {
      const result = new Set();
      for (const item of this) {
        if (other.has(item)) result.add(item);
      }
      return result;
    },
    writable: true,
    configurable: true,
  });
}
`;

new Function(RUNTIME_POLYFILLS)();

const WORKER_FILE_CONTENT = RUNTIME_POLYFILLS + __PDFJS_WORKER_SOURCE__;

const execFileAsync = promisify(execFile);

const SOFFICE_CANDIDATES = [
  "/Applications/LibreOffice.app/Contents/MacOS/soffice",
  "/opt/homebrew/bin/soffice",
  "/usr/local/bin/soffice",
  "/usr/bin/soffice",
  "/usr/bin/libreoffice",
];

export async function findSoffice(): Promise<string | null> {
  for (const p of SOFFICE_CANDIDATES) {
    try {
      await fs.access(p);
      return p;
    } catch {
      // try next
    }
  }
  try {
    const { stdout } = await execFileAsync("/bin/sh", [
      "-c",
      "command -v soffice 2>/dev/null || command -v libreoffice 2>/dev/null",
    ]);
    const trimmed = stdout.trim();
    if (trimmed) return trimmed;
  } catch {
    // fall through
  }
  return null;
}

// Kept outside the vault so iCloud Drive doesn't sync per-workbook PDFs across
// machines. Content-hashed names make collisions safe and re-creation cheap.
export function pluginCacheDir(pluginId: string): string {
  return path.join(os.tmpdir(), `${pluginId}-cache`);
}

export async function convertOfficeToPdf(
  sofficeBin: string,
  buf: ArrayBuffer,
  originalExtension: string,
  cacheDir: string,
): Promise<string> {
  const bytes = new Uint8Array(buf);
  const hash = createHash("sha256").update(bytes).digest("hex").slice(0, 16);
  await fs.mkdir(cacheDir, { recursive: true });

  const cachedPdf = path.join(cacheDir, `${hash}.pdf`);
  try {
    await fs.access(cachedPdf);
    return cachedPdf;
  } catch {
    // cache miss; convert
  }

  // soffice names the output `<input-basename>.pdf` in the outdir, so we
  // stage the input under a hashed basename to make the resulting path
  // predictable and unique per content hash.
  const stagedInput = path.join(cacheDir, `${hash}.${originalExtension}`);
  await fs.writeFile(stagedInput, bytes);
  try {
    await execFileAsync(
      sofficeBin,
      ["--headless", "--convert-to", "pdf", "--outdir", cacheDir, stagedInput],
      { timeout: 120000 },
    );
    await fs.access(cachedPdf);
    return cachedPdf;
  } finally {
    try {
      await fs.unlink(stagedInput);
    } catch {
      // best-effort cleanup
    }
  }
}

// Write the bundled worker source to a real file inside the plugin install
// dir, then hand PDF.js an `app://`-scheme URL via the vault adapter. This
// avoids loading the worker from a blob: URL — under blob:, `import.meta.url`
// inside the worker resolves to the blob URL itself, and PDF.js v5's asset
// resolution (wasm, cmaps, standard fonts) breaks.
let pdfjsWorkerInitialized = false;

export async function ensurePdfjsWorker(app: App, pluginId: string): Promise<void> {
  if (pdfjsWorkerInitialized) return;
  const adapter = app.vault.adapter;
  if (!(adapter instanceof FileSystemAdapter)) {
    throw new Error("PDF.js setup requires a local-disk vault.");
  }
  const workerVaultPath = `${app.vault.configDir}/plugins/${pluginId}/pdf.worker.mjs`;
  let needsWrite = true;
  try {
    const existing = await adapter.read(workerVaultPath);
    if (existing === WORKER_FILE_CONTENT) needsWrite = false;
  } catch {
    // missing — will write below
  }
  if (needsWrite) {
    await adapter.write(workerVaultPath, WORKER_FILE_CONTENT);
  }
  pdfjsLib.GlobalWorkerOptions.workerSrc = adapter.getResourcePath(workerVaultPath);
  pdfjsWorkerInitialized = true;
}

export interface RenderedPdfPageInfo {
  pageNumber: number;
  width: number;
  height: number;
  failed: boolean;
}

export interface RenderOptions {
  scale?: number;
  isStale?: () => boolean;
}

export async function renderPdfPagesIntoStage(
  pdfPath: string,
  stage: HTMLElement,
  slideClass: string,
  canvasClass: string,
  opts: RenderOptions = {},
): Promise<RenderedPdfPageInfo[]> {
  const { scale = 2, isStale } = opts;
  const pdfBytes = await fs.readFile(pdfPath);
  // useWasm / OffscreenCanvas / ImageDecoder all try to resolve sibling assets
  // via import.meta.url, which is the worker's blob: URL here and doesn't
  // resolve. For LibreOffice-produced PDFs none of these paths add value, so
  // disable them to keep the worker on the pure-JS rendering path.
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(pdfBytes),
    useWasm: false,
    isOffscreenCanvasSupported: false,
    isImageDecoderSupported: false,
    // Draw glyphs as canvas paths instead of going through @font-face. The
    // @font-face route mis-renders some characters from LibreOffice's subset
    // fonts (whole letters drop out). Path rendering uses the embedded
    // outlines directly and is robust to subset quirks.
    disableFontFace: true,
  });
  const pdf = await loadingTask.promise;

  const out: RenderedPdfPageInfo[] = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    if (isStale?.()) break;
    const wrap = stage.createDiv({ cls: slideClass });
    const canvas = wrap.createEl("canvas", { cls: canvasClass });
    try {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale });
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("2d canvas context unavailable");
      await page.render({ canvasContext: ctx, viewport, canvas }).promise;
      out.push({
        pageNumber: i,
        width: viewport.width,
        height: viewport.height,
        failed: false,
      });
    } catch (e) {
      console.error(`PDF render failed on page ${i}:`, e);
      canvas.remove();
      wrap
        .createDiv({ cls: "docx-claude-pdf-error" })
        .setText(
          `Page ${i} failed to render: ${e instanceof Error ? e.message : String(e)}`,
        );
      out.push({ pageNumber: i, width: 0, height: 0, failed: true });
    }
  }
  return out;
}
