import { Notice, TFile } from "obsidian";
import { PPTXViewer } from "pptxviewjs";
import JSZip from "jszip";
import { OfficeFileView } from "./OfficeFileView";

export const PPTX_CLAUDE_VIEW_TYPE = "pptx-claude-view";

// PowerPoint widescreen default is 12192000 x 6858000 EMU = 13.33" x 7.5"
// = 1920 x 1080 at 144 DPI. Render at retina (2x) sharpness so detail holds
// up when the user scales/zooms.
const SLIDE_WIDTH = 1920;
const SLIDE_HEIGHT = 1080;

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
    const stage = this.renderEl.createDiv({ cls: "docx-claude-pptx-stage" });
    const buf = await this.app.vault.readBinary(file);

    // Read the zip ourselves so we always know the true slide list, even if
    // pptxviewjs's loader bails out partway through.
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
      // eslint-disable-next-line no-console
      console.error("pptxviewjs load failed; falling back to text-only rendering:", e);
      viewer = null;
    }

    const visualSuccess: number[] = [];
    const textFallback: number[] = [];

    for (let i = 0; i < totalSlides; i++) {
      const oneIdx = i + 1;
      const hasViewerSlide = viewer && i < viewerSlideCount;
      let renderedVisually = false;

      if (hasViewerSlide) {
        const canvasWrap = stage.createDiv({ cls: "docx-claude-pptx-slide" });
        const canvas = canvasWrap.createEl("canvas", {
          cls: "docx-claude-pptx-canvas",
        });
        canvas.width = SLIDE_WIDTH;
        canvas.height = SLIDE_HEIGHT;
        try {
          await viewer!.renderSlide(i, canvas, { quality: "high" });
          renderedVisually = true;
          visualSuccess.push(oneIdx);
        } catch (e) {
          // eslint-disable-next-line no-console
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
          // eslint-disable-next-line no-console
          console.error(`Text fallback failed on slide ${oneIdx}:`, e);
        }
      }
    }

    if (viewer) viewer.destroy();

    const summary = this.buildSummary(totalSlides, visualSuccess.length, textFallback.length);
    if (summary) new Notice(summary, 6000);
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

  private buildSummary(total: number, visual: number, fallback: number): string | null {
    if (visual === total) return null;
    const parts = [`Rendered ${visual} of ${total} slides visually.`];
    if (fallback > 0) parts.push(`${fallback} shown as text only.`);
    parts.push("Use Open in PowerPoint for the full deck.");
    return parts.join(" ");
  }
}

function slideIndex(path: string): number {
  const m = path.match(/slide(\d+)\.xml$/);
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
