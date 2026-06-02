import * as pdfjsLib from "pdfjs-dist";
import { App, FileSystemAdapter } from "obsidian";
import { promises as fs } from "node:fs";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { createHash } from "node:crypto";
import * as path from "node:path";
import * as os from "node:os";
import { pathToFileURL } from "node:url";

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

// Bump when the conversion command/flags change (or when a font-setup change
// alters substitution) so cached PDFs produced under older settings are
// invalidated instead of silently reused. Folded into the cache key.
const CONVERSION_VERSION = "4";

// PDF export FilterData shared across the Writer/Calc/Impress export filters.
// UseLosslessCompression avoids JPEG artifacts on embedded images and
// ReduceImageResolution:false keeps images at full resolution — both maximize
// visual fidelity for a static on-screen viewer (at the cost of larger PDFs).
const PDF_FILTER_DATA =
  '{"UseLosslessCompression":{"type":"boolean","value":"true"},' +
  '"ReduceImageResolution":{"type":"boolean","value":"false"}}';

// Calc-specific FilterData. SinglePageSheets:true renders one un-paginated page
// per sheet sized to the full used range, ignoring the workbook's print area
// and paper size. Without it Calc honors the print area and clips to it (and
// paginates wide/long sheets), which surfaces as a cut-off preview.
const CALC_PDF_FILTER_DATA =
  '{"SinglePageSheets":{"type":"boolean","value":"true"},' +
  '"UseLosslessCompression":{"type":"boolean","value":"true"},' +
  '"ReduceImageResolution":{"type":"boolean","value":"false"}}';

// Map a source extension to LibreOffice's app-specific PDF export filter so the
// FilterData options above are honored. Unknown types fall back to plain "pdf".
function pdfFilterFor(ext: string): string {
  const e = ext.toLowerCase();
  if (["xls", "xlsx", "xlsm", "csv", "ods"].includes(e)) {
    return `pdf:calc_pdf_Export:${CALC_PDF_FILTER_DATA}`;
  }
  if (["ppt", "pptx", "odp"].includes(e)) {
    return `pdf:impress_pdf_Export:${PDF_FILTER_DATA}`;
  }
  if (["doc", "docx", "rtf", "odt", "txt"].includes(e)) {
    return `pdf:writer_pdf_Export:${PDF_FILTER_DATA}`;
  }
  return "pdf";
}

// Office fonts absent on macOS that have installed equivalents. Carlito/Caladea
// are metric-compatible with Calibri/Cambria. "Aptos Display" is only the
// font's legacy family name (its typographic family is "Aptos"), so macOS
// CoreText can't match it and LibreOffice falls back to a serif; mapping it to
// "Aptos" lets LibreOffice resolve the Aptos family and pick the Display face.
const FONT_SUBSTITUTIONS: ReadonlyArray<readonly [string, string]> = [
  ["Calibri", "Carlito"],
  ["Calibri Light", "Carlito"],
  ["Cambria", "Caladea"],
  ["Aptos Display", "Aptos"],
];

function buildSubstitutionXcu(): string {
  const nodes = FONT_SUBSTITUTIONS.map(
    ([from, to], i) =>
      `  <node oor:name="sub${i}" oor:op="replace">\n` +
      `   <prop oor:name="ReplaceFont" oor:op="fuse"><value>${from}</value></prop>\n` +
      `   <prop oor:name="SubstituteFont" oor:op="fuse"><value>${to}</value></prop>\n` +
      `   <prop oor:name="OnScreenOnly" oor:op="fuse"><value>false</value></prop>\n` +
      `   <prop oor:name="Always" oor:op="fuse"><value>true</value></prop>\n` +
      `  </node>`,
  ).join("\n");
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n` +
    ` <item oor:path="/org.openoffice.Office.Common/Font/Substitution">\n` +
    `  <prop oor:name="Replacement" oor:op="fuse"><value>true</value></prop>\n` +
    ` </item>\n` +
    ` <item oor:path="/org.openoffice.Office.Common/Font/Substitution/FontPairs">\n` +
    `${nodes}\n` +
    ` </item>\n` +
    `</oor:items>\n`
  );
}

// A dedicated LibreOffice user profile, separate from any interactive install,
// so the font substitutions are seeded without touching the user's own profile
// and conversions don't fight a running LibreOffice for the profile lock.
// Seeded once per CONVERSION_VERSION (LibreOffice persists it across runs).
let sofficeProfileInit: Promise<string> | null = null;

function resetSofficeProfileInit(): void {
  sofficeProfileInit = null;
}

async function ensureSofficeProfile(cacheDir: string): Promise<string> {
  if (!sofficeProfileInit) {
    sofficeProfileInit = (async () => {
      const profileDir = `${cacheDir}-loprofile`;
      const userDir = path.join(profileDir, "user");
      const marker = path.join(userDir, `.fontsub-${CONVERSION_VERSION}`);
      try {
        await fs.access(marker);
        return profileDir;
      } catch {
        // not yet seeded for this CONVERSION_VERSION
      }
      await fs.mkdir(userDir, { recursive: true });
      await fs.writeFile(
        path.join(userDir, "registrymodifications.xcu"),
        buildSubstitutionXcu(),
        "utf-8",
      );
      await fs.writeFile(marker, "", "utf-8");
      return profileDir;
    })().catch((err) => {
      resetSofficeProfileInit();
      throw err;
    });
  }
  return sofficeProfileInit;
}

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

// Coalesce concurrent conversions of identical content so two views opening
// the same file don't both write the staged input + invoke soffice + race on
// the output path.
const inFlightConversions = new Map<string, Promise<string>>();

export async function convertOfficeToPdf(
  sofficeBin: string,
  buf: ArrayBuffer,
  originalExtension: string,
  cacheDir: string,
): Promise<string> {
  const bytes = new Uint8Array(buf);
  const hash = createHash("sha256")
    .update(bytes)
    .update(` ${CONVERSION_VERSION} ${originalExtension.toLowerCase()}`)
    .digest("hex")
    .slice(0, 16);
  await fs.mkdir(cacheDir, { recursive: true });

  const cachedPdf = path.join(cacheDir, `${hash}.pdf`);
  try {
    await fs.access(cachedPdf);
    return cachedPdf;
  } catch {
    // cache miss; convert
  }

  const existing = inFlightConversions.get(hash);
  if (existing) return existing;

  const work = (async () => {
    // soffice names the output `<input-basename>.pdf` in the outdir, so we
    // stage under a hashed basename to make the result path predictable and
    // unique per content hash.
    const stagedInput = path.join(cacheDir, `${hash}.${originalExtension}`);
    await fs.writeFile(stagedInput, bytes);
    const profileDir = await ensureSofficeProfile(cacheDir);
    try {
      await execFileAsync(
        sofficeBin,
        [
          "--headless",
          `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
          "--convert-to",
          pdfFilterFor(originalExtension),
          "--outdir",
          cacheDir,
          stagedInput,
        ],
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
  })();

  inFlightConversions.set(hash, work);
  try {
    return await work;
  } finally {
    inFlightConversions.delete(hash);
  }
}

// Write the bundled worker source to a real file inside the plugin install
// dir, then hand PDF.js an `app://`-scheme URL via the vault adapter. This
// avoids loading the worker from a blob: URL — under blob:, `import.meta.url`
// inside the worker resolves to the blob URL itself, and PDF.js v5's asset
// resolution (wasm, cmaps, standard fonts) breaks.
let pdfjsWorkerInit: Promise<void> | null = null;

export function ensurePdfjsWorker(app: App, pluginId: string): Promise<void> {
  // Memoize the in-flight init so two concurrent first-time callers can't
  // both pass the existence check and race the file write.
  if (!pdfjsWorkerInit) {
    pdfjsWorkerInit = doEnsurePdfjsWorker(app, pluginId).catch((err) => {
      pdfjsWorkerInit = null;
      throw err;
    });
  }
  return pdfjsWorkerInit;
}

async function doEnsurePdfjsWorker(app: App, pluginId: string): Promise<void> {
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
}

export interface PdfReadyInfo {
  numPages: number;
  firstPageFailed: boolean;
}

export interface RenderOptions {
  isStale?: () => boolean;
}

export interface PdfRenderHandle {
  // Resolves once the document is loaded and the first page has been rendered,
  // so the caller can decide whether to keep the PDF view or fall back. The
  // remaining pages render lazily as they scroll into view.
  ready: Promise<PdfReadyInfo>;
  cancel(): void;
}

export function renderPdfPagesIntoStage(
  pdfPath: string,
  stage: HTMLElement,
  slideClass: string,
  canvasClass: string,
  opts: RenderOptions = {},
): PdfRenderHandle {
  const { isStale } = opts;
  let cancelled = false;
  let loadingTask: ReturnType<typeof pdfjsLib.getDocument> | null = null;
  let currentRenderTask: { cancel?: () => void } | null = null;
  let observer: IntersectionObserver | null = null;
  const stale = (): boolean => cancelled || isStale?.() === true;

  // Crisp on HiDPI: size each canvas bitmap to device pixels for the CSS max
  // display width, then let CSS scale it down to fit the pane. devicePixelRatio
  // is capped at 2 so memory stays bounded on very high-DPR displays.
  const dpr = Math.min(Math.max(window.devicePixelRatio || 1, 1), 2);
  const TARGET_CSS_WIDTH = 1280; // matches .docx-claude-pdf-canvas max-width
  const MAX_CANVAS_DIM = 8192; // keep within browser/GPU canvas size limits
  const MAX_RENDERED = 8; // LRU cap on simultaneously-rasterized pages

  const ready = (async (): Promise<PdfReadyInfo> => {
    const pdfBytes = await fs.readFile(pdfPath);
    if (stale()) throw new Error("PDF render cancelled");
    // useWasm / OffscreenCanvas / ImageDecoder all try to resolve sibling
    // assets via import.meta.url, which is the worker's blob: URL here and
    // doesn't resolve. For LibreOffice-produced PDFs none of these paths add
    // value, so disable them to keep the worker on pure-JS rendering.
    loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(pdfBytes),
      useWasm: false,
      isOffscreenCanvasSupported: false,
      isImageDecoderSupported: false,
      // Draw glyphs as canvas paths instead of going through @font-face. The
      // @font-face route mis-renders LibreOffice's subset fonts.
      disableFontFace: true,
    });
    const pdf = await loadingTask.promise;
    if (stale()) {
      try { await loadingTask.destroy(); } catch { /* ignore */ }
      throw new Error("PDF render cancelled");
    }

    const numPages = pdf.numPages;
    if (numPages === 0) return { numPages: 0, firstPageFailed: true };

    interface PageSlot {
      index: number;
      wrap: HTMLElement;
      state: "pending" | "rendering" | "rendered" | "failed";
    }
    const slots: PageSlot[] = [];
    const renderedOrder: number[] = []; // page indices currently rasterized

    // Placeholders for every page so total scroll height is correct from the
    // start; sized to page 1's aspect ratio and corrected per page on render.
    const firstViewport = (await pdf.getPage(1)).getViewport({ scale: 1 });
    const defaultRatio = `${firstViewport.width} / ${firstViewport.height}`;
    for (let i = 1; i <= numPages; i++) {
      const wrap = stage.createDiv({ cls: slideClass });
      wrap.dataset.pageIndex = String(i);
      wrap.style.aspectRatio = defaultRatio;
      slots.push({ index: i, wrap, state: "pending" });
    }

    // Drop the canvas of whichever rendered page is furthest (by index, which
    // tracks scroll distance) from the page just rendered, keeping peak memory
    // bounded. Evicted pages re-render when scrolled back into view.
    const evictIfNeeded = (anchor: number): void => {
      while (renderedOrder.length > MAX_RENDERED) {
        let pos = 0;
        let far = -1;
        for (let k = 0; k < renderedOrder.length; k++) {
          const d = Math.abs(renderedOrder[k] - anchor);
          if (d > far) { far = d; pos = k; }
        }
        const idx = renderedOrder.splice(pos, 1)[0];
        const s = slots[idx - 1];
        if (s && s.state === "rendered") {
          s.wrap.querySelector("canvas")?.remove();
          s.state = "pending";
        }
      }
    };

    const renderSlot = async (slot: PageSlot): Promise<void> => {
      const canvas = slot.wrap.createEl("canvas", { cls: canvasClass });
      try {
        const page = await pdf.getPage(slot.index);
        if (stale()) { canvas.remove(); slot.state = "pending"; return; }
        const base = page.getViewport({ scale: 1 });
        slot.wrap.style.aspectRatio = `${base.width} / ${base.height}`;
        let scale = (TARGET_CSS_WIDTH * dpr) / base.width;
        if (scale < 1) scale = 1;
        let viewport = page.getViewport({ scale });
        const maxDim = Math.max(viewport.width, viewport.height);
        if (maxDim > MAX_CANVAS_DIM) {
          viewport = page.getViewport({ scale: scale * (MAX_CANVAS_DIM / maxDim) });
        }
        canvas.width = Math.floor(viewport.width);
        canvas.height = Math.floor(viewport.height);
        const ctx = canvas.getContext("2d");
        if (!ctx) throw new Error("2d canvas context unavailable");
        const task = page.render({ canvasContext: ctx, viewport, canvas });
        currentRenderTask = task as unknown as { cancel?: () => void };
        await task.promise;
        currentRenderTask = null;
        try { page.cleanup(); } catch { /* ignore */ }
        if (stale()) { canvas.remove(); slot.state = "pending"; return; }
        slot.state = "rendered";
        renderedOrder.push(slot.index);
        evictIfNeeded(slot.index);
      } catch (e) {
        currentRenderTask = null;
        if (stale()) { canvas.remove(); slot.state = "pending"; return; }
        console.error(`PDF render failed on page ${slot.index}:`, e);
        canvas.remove();
        slot.wrap
          .createDiv({ cls: "docx-claude-pdf-error" })
          .setText(
            `Page ${slot.index} failed to render: ${e instanceof Error ? e.message : String(e)}`,
          );
        slot.state = "failed";
      }
    };

    // Serialize lazy renders so two pdf.js render tasks never run concurrently.
    let chain: Promise<void> = Promise.resolve();
    const queueRender = (slot: PageSlot): void => {
      if (slot.state !== "pending") return;
      slot.state = "rendering";
      chain = chain.then(() => (stale() ? undefined : renderSlot(slot)));
    };

    // Render the first page eagerly: gives the caller a success/failure signal
    // and shows content immediately. (No observer is active yet, so the slot's
    // "pending" state is fine; renderSlot drives it to rendered/failed.)
    await renderSlot(slots[0]);
    const firstPageFailed = slots[0].state === "failed";

    if (!stale() && numPages > 1) {
      const root = stage.closest(".docx-claude-scroll") as HTMLElement | null;
      observer = new IntersectionObserver(
        (entries) => {
          if (stale()) return;
          for (const entry of entries) {
            if (!entry.isIntersecting) continue;
            const a = (entry.target as HTMLElement).dataset.pageIndex;
            if (!a) continue;
            const slot = slots[parseInt(a, 10) - 1];
            if (slot && slot.state === "pending") queueRender(slot);
          }
        },
        { root, rootMargin: "150% 0px", threshold: 0 },
      );
      for (const slot of slots) observer.observe(slot.wrap);
    }

    return { numPages, firstPageFailed };
  })();

  return {
    ready,
    cancel(): void {
      cancelled = true;
      if (observer) { observer.disconnect(); observer = null; }
      try { currentRenderTask?.cancel?.(); } catch { /* ignore */ }
      if (loadingTask) {
        loadingTask.destroy().catch(() => { /* ignore */ });
      }
    },
  };
}
