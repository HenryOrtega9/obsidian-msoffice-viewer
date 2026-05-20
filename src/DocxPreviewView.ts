import { TFile } from "obsidian";
import { renderAsync } from "docx-preview";
import { OfficeFileView } from "./OfficeFileView";

export const DOCX_CLAUDE_VIEW_TYPE = "docx-claude-view";

export class DocxPreviewView extends OfficeFileView {
  getViewType(): string {
    return DOCX_CLAUDE_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Docx preview";
  }

  getIcon(): string {
    return "file-text";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "docx";
  }

  protected getExternalAppLabel(): string {
    return "Open in Word";
  }

  protected async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.renderEl.empty();
    const buf = await this.app.vault.readBinary(file);
    await renderAsync(buf, this.renderEl, this.renderEl, {
      className: "docx-claude",
      ignoreLastRenderedPageBreak: true,
      experimental: true,
    });
  }
}
