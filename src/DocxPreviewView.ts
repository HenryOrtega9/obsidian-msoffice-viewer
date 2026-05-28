import { Notice, TFile } from "obsidian";
import { renderAsync } from "docx-preview";
import { OfficeFileView } from "./OfficeFileView";
import { findSoffice } from "./officeToPdf";
import { buildDocxOptions } from "./docx/options";
import { isRenderEmpty } from "./docx/emptyRender";
import { warn } from "./docx/warn";

export const DOCX_CLAUDE_VIEW_TYPE = "docx-claude-view";

type RenderMode = "html" | "pdf";

export class DocxPreviewView extends OfficeFileView {
  private sofficeAvailable = false;
  private renderMode: RenderMode = "html";
  // Sticky manual override; null means auto-route by content. Reset on file
  // switch (onLoadFile) but preserved across toggle-triggered re-renders.
  private userForcedMode: RenderMode | null = null;
  private toggleBtn: HTMLButtonElement | null = null;

  async onLoadFile(file: TFile): Promise<void> {
    this.userForcedMode = null;
    await super.onLoadFile(file);
  }

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
    if (this.file !== file || !this.renderEl) return;

    // LibreOffice → PDF is the high-fidelity default: it renders through the
    // real Office layout engine. docx-preview is the fallback when LibreOffice
    // is unavailable or fails, and stays reachable via the toolbar toggle.
    const mode = this.userForcedMode ?? "pdf";

    if (mode === "pdf") {
      if (await this.tryRenderPdf(file)) return;
      if (this.file !== file || !this.renderEl) return;
      this.renderEl.empty();
      if (await this.tryRenderHtml(buf)) return;
      this.showError(!this.sofficeAvailable);
      return;
    }

    // Manual override to the selectable text renderer; PDF rescues if it fails.
    if (await this.tryRenderHtml(buf)) return;
    if (this.file !== file || !this.renderEl) return;
    this.renderEl.empty();
    if (await this.tryRenderPdf(file)) return;
    this.showError(!this.sofficeAvailable);
  }

  // Render with docx-preview. Returns false (so the caller can fall back) when
  // it throws or produces an empty wrapper.
  private async tryRenderHtml(buf: ArrayBuffer): Promise<boolean> {
    if (!this.renderEl) return false;
    try {
      await renderAsync(buf, this.renderEl, this.renderEl, buildDocxOptions());
      if (!this.renderEl) return false;
      if (isRenderEmpty(this.renderEl)) {
        warn("render-empty", null, { file: this.file?.path });
        return false;
      }
      this.renderMode = "html";
      this.updateToggleLabel();
      return true;
    } catch (e) {
      warn("render", e, { file: this.file?.path });
      return false;
    }
  }

  // Render via the shared LibreOffice -> PDF -> PDF.js pipeline. Returns false
  // when LibreOffice is missing or the conversion fails.
  private async tryRenderPdf(file: TFile): Promise<boolean> {
    if (!this.renderEl) return false;
    const soffice = await findSoffice();
    if (this.file !== file || !this.renderEl) return false;
    this.sofficeAvailable = soffice != null;
    if (!soffice) return false;
    this.renderEl.empty();
    try {
      await this.renderViaLibreOfficePdf(file, soffice, "docx");
      this.renderMode = "pdf";
      this.updateToggleLabel();
      return true;
    } catch (e) {
      warn("libreoffice", e, { file: file.path });
      return false;
    }
  }

  protected buildExtraToolbar(toolbar: HTMLElement): void {
    this.toggleBtn = toolbar.createEl("button", {
      cls: "docx-claude-zoom-btn docx-claude-fidelity-toggle",
      text: this.toggleLabel(),
      attr: { title: this.toggleTitle(), "aria-label": this.toggleTitle() },
    });
    this.toggleBtn.addEventListener("click", (ev) => {
      ev.preventDefault();
      void this.onToggleClick();
    });
  }

  private toggleLabel(): string {
    return this.renderMode === "pdf" ? "Back to text" : "High fidelity (PDF)";
  }

  private toggleTitle(): string {
    return this.renderMode === "pdf"
      ? "Switch back to the selectable text renderer"
      : "Re-render this document via LibreOffice for higher fidelity";
  }

  private updateToggleLabel(): void {
    if (!this.toggleBtn) return;
    this.toggleBtn.setText(this.toggleLabel());
    this.toggleBtn.setAttribute("title", this.toggleTitle());
    this.toggleBtn.setAttribute("aria-label", this.toggleTitle());
  }

  private async onToggleClick(): Promise<void> {
    if (!this.file) return;
    const target: RenderMode = this.renderMode === "pdf" ? "html" : "pdf";
    if (target === "pdf") {
      const soffice = await findSoffice();
      if (!soffice) {
        new Notice("LibreOffice not found. Install it to use high-fidelity PDF rendering.", 6000);
        return;
      }
    }
    this.userForcedMode = target;
    await this.renderFile(this.file);
  }

  private showError(noLibreOffice: boolean): void {
    if (!this.renderEl) return;
    this.renderEl.empty();
    this.renderEl
      .createDiv({ cls: "docx-claude-pdf-error" })
      .setText(
        noLibreOffice
          ? "Couldn't render this .docx, and the high-fidelity fallback needs LibreOffice. Install LibreOffice or use Open in Word."
          : "Couldn't render this .docx. See console for details.",
      );
    new Notice("Failed to render .docx. See console for details.", 6000);
  }
}
