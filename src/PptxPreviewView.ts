import { TFile } from "obsidian";
import { init } from "pptx-preview";
import { OfficeFileView } from "./OfficeFileView";

export const PPTX_CLAUDE_VIEW_TYPE = "pptx-claude-view";

const SLIDE_WIDTH = 960;
const SLIDE_HEIGHT = 540;

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
    const viewer = init(stage, {
      width: SLIDE_WIDTH,
      height: SLIDE_HEIGHT,
      mode: "list",
    });
    await viewer.preview(buf);
  }
}
