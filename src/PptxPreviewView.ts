import { Notice, TFile } from "obsidian";
import { PPTXViewer } from "pptxviewjs";
import JSZip from "jszip";
import type { Chart } from "chart.js";
import { OfficeFileView } from "./OfficeFileView";
import { findSoffice } from "./officeToPdf";
import { renderDeckNatively } from "./pptx/render";

export const PPTX_CLAUDE_VIEW_TYPE = "pptx-claude-view";

const FALLBACK_SLIDE_WIDTH = 1920;
const FALLBACK_SLIDE_HEIGHT = 1080;

export class PptxPreviewView extends OfficeFileView {
  private imageObjectUrls: string[] = [];
  private chartInstances: Chart[] = [];

  getViewType(): string {
    return PPTX_CLAUDE_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Pptx preview";
  }

  getIcon(): string {
    return "presentation";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "pptx";
  }

  protected getExternalAppLabel(): string {
    return "Open in PowerPoint";
  }

  async onUnloadFile(file: TFile): Promise<void> {
    this.resetState();
    await super.onUnloadFile(file);
  }

  private resetState(): void {
    this.revokeImageUrls();
    this.destroyCharts();
    this.setEngineLabel("");
  }

  private revokeImageUrls(): void {
    for (const url of this.imageObjectUrls) URL.revokeObjectURL(url);
    this.imageObjectUrls = [];
  }

  private destroyCharts(): void {
    for (const chart of this.chartInstances) {
      try { chart.destroy(); } catch { /* ignore */ }
    }
    this.chartInstances = [];
  }

  // Fallback chain: LibreOffice PDF (primary, high fidelity) -> native OOXML
  // renderer -> pptxviewjs canvas -> raw-XML text card. Each tier reports its
  // engine via the toolbar badge.
  protected async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.resetState();
    this.renderEl.empty();

    const sofficeBin = await findSoffice();
    if (sofficeBin) {
      try {
        await this.renderViaLibreOfficePdf(file, sofficeBin, "pptx");
        // renderViaLibreOfficePdf returns silently when stale; don't stamp the
        // engine badge onto whatever file the view shows now.
        if (this.file !== file) return;
        this.setEngineLabel("LibreOffice");
        return;
      } catch (e) {
        console.error("LibreOffice render failed; trying native renderer:", e);
      }
    }

    if (this.file !== file || !this.renderEl) return;
    this.resetState();
    this.renderEl.empty();
    try {
      await this.renderViaNative(file);
      if (this.file !== file) return;
      this.setEngineLabel("Native");
      return;
    } catch (e) {
      console.error("Native PPTX render failed; falling back to pptxviewjs:", e);
    }

    if (this.file !== file || !this.renderEl) return;
    this.resetState();
    this.renderEl.empty();
    if (!sofficeBin) {
      new Notice(
        "LibreOffice not found; rendering with the built-in engine (lower fidelity).",
        4000,
      );
    }
    this.setEngineLabel("Built-in");
    try {
      await this.renderViaPptxViewJS(file);
    } catch (e) {
      // Last tier: surface the failure instead of leaving a blank pane with an
      // unhandled rejection (e.g. a corrupt zip fails JSZip.loadAsync).
      console.error("pptxviewjs fallback failed:", e);
      if (this.file !== file || !this.renderEl) return;
      this.renderEl.empty();
      this.renderEl.createDiv({
        cls: "docx-claude-pdf-error",
        text: "Could not render this .pptx file. It may be corrupt. Use Open in PowerPoint instead.",
      });
    }
  }

  private async renderViaNative(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    const buf = await this.app.vault.readBinary(file);
    if (this.file !== file || !this.renderEl) return;
    const result = await renderDeckNatively(buf, this.renderEl, {
      isStale: () => this.file !== file,
    });
    if (this.file !== file) {
      for (const url of result.objectUrls) URL.revokeObjectURL(url);
      for (const chart of result.charts) {
        try { chart.destroy(); } catch { /* ignore */ }
      }
      return;
    }
    this.imageObjectUrls = result.objectUrls;
    this.chartInstances = result.charts;
  }

  private async renderViaPptxViewJS(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    const stage = this.renderEl.createDiv({ cls: "docx-claude-pdf-stage" });
    const buf = await this.app.vault.readBinary(file);

    const zip = await JSZip.loadAsync(buf);
    const slidePaths: string[] = [];
    zip.forEach((relPath) => {
      if (/^ppt\/slides\/slide\d+\.xml$/.test(relPath)) slidePaths.push(relPath);
    });
    slidePaths.sort((a, b) => slideIndex(a) - slideIndex(b));
    const totalSlides = slidePaths.length;
    if (totalSlides === 0) {
      new Notice("No slides found in this .pptx file.");
      return;
    }

    let viewer: PPTXViewer | null = null;
    let viewerSlideCount = 0;
    try {
      viewer = new PPTXViewer({ slideSizeMode: "fit", backgroundColor: "#ffffff" });
      await viewer.loadFile(buf);
      viewerSlideCount = viewer.getSlideCount();
    } catch (e) {
      console.error("pptxviewjs load failed:", e);
      viewer = null;
    }

    const visualSuccess: number[] = [];
    const textFallback: number[] = [];

    try {
      for (let i = 0; i < totalSlides; i++) {
        if (this.file !== file) break;
        const oneIdx = i + 1;
        let renderedVisually = false;
        if (viewer && i < viewerSlideCount) {
          const canvasWrap = stage.createDiv({ cls: "docx-claude-pdf-slide" });
          const canvas = canvasWrap.createEl("canvas", { cls: "docx-claude-pdf-canvas" });
          canvas.width = FALLBACK_SLIDE_WIDTH;
          canvas.height = FALLBACK_SLIDE_HEIGHT;
          try {
            await viewer.renderSlide(i, canvas, { quality: "high" });
            renderedVisually = true;
            visualSuccess.push(oneIdx);
          } catch (e) {
            console.error(`pptxviewjs render failed on slide ${oneIdx}:`, e);
            canvasWrap.remove();
          }
        }
        if (!renderedVisually) {
          try {
            const xml = await zip.file(slidePaths[i])!.async("text");
            this.renderTextFallback(stage, oneIdx, xml);
            textFallback.push(oneIdx);
          } catch (e) {
            console.error(`Text fallback failed on slide ${oneIdx}:`, e);
          }
        }
      }
    } finally {
      viewer?.destroy();
    }

    if (visualSuccess.length < totalSlides) {
      const parts = [`Rendered ${visualSuccess.length} of ${totalSlides} slides visually.`];
      if (textFallback.length > 0) parts.push(`${textFallback.length} shown as text only.`);
      parts.push("Use Open in PowerPoint for the full deck.");
      new Notice(parts.join(" "), 6000);
    }
  }

  private renderTextFallback(parent: HTMLElement, oneBasedIndex: number, xml: string): void {
    const card = parent.createDiv({ cls: "docx-claude-pptx-fallback" });
    card.createEl("div", {
      cls: "docx-claude-pptx-fallback-title",
      text: `Slide ${oneBasedIndex} — text only`,
    });
    const body = card.createDiv({ cls: "docx-claude-pptx-fallback-body" });
    const runs = extractTextRuns(xml);
    if (runs.length === 0) {
      body.createEl("p", {
        text: "(no readable text in this slide)",
        cls: "docx-claude-pptx-fallback-empty",
      });
      return;
    }
    for (const run of runs) body.createEl("p", { text: run });
  }
}

function slideIndex(p: string): number {
  const m = p.match(/slide(\d+)\.xml$/);
  return m ? parseInt(m[1], 10) : 0;
}

function extractTextRuns(xml: string): string[] {
  const out: string[] = [];
  const tagRe = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
  let m: RegExpExecArray | null;
  while ((m = tagRe.exec(xml)) !== null) {
    const decoded = decodeXmlEntities(m[1]);
    if (decoded.trim()) out.push(decoded);
  }
  return out;
}

function decodeXmlEntities(s: string): string {
  // &amp; must decode LAST or "&amp;lt;" double-decodes to "<".
  // fromCodePoint handles astral chars (emoji) that fromCharCode corrupts.
  return s
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
