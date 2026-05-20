import { Notice, TFile } from "obsidian";
import { PPTXViewer } from "pptxviewjs";
import JSZip from "jszip";
import { OfficeFileView } from "./OfficeFileView";
import { findSoffice } from "./officeToPdf";

export const PPTX_CLAUDE_VIEW_TYPE = "pptx-claude-view";

const FALLBACK_SLIDE_WIDTH = 1920;
const FALLBACK_SLIDE_HEIGHT = 1080;

export class PptxPreviewView extends OfficeFileView {
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

  protected async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.renderEl.empty();

    const sofficeBin = await findSoffice();
    if (sofficeBin) {
      try {
        await this.renderViaLibreOfficePdf(file, sofficeBin, "pptx");
        return;
      } catch (e) {
        console.error("LibreOffice render failed; falling back to pptxviewjs:", e);
        new Notice("LibreOffice rendering failed; using fallback renderer.");
      }
    } else {
      new Notice(
        "LibreOffice not found; rendering with the JS fallback (lower fidelity).",
        4000,
      );
    }

    if (this.file !== file || !this.renderEl) return;
    this.renderEl.empty();
    await this.renderViaPptxViewJS(file);
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
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
}
