import { FileView, TFile, WorkspaceLeaf } from "obsidian";
import { renderAsync } from "docx-preview";
import { computeLocator } from "./selection";
import type { StructuralLocator } from "./types";

export const DOCX_CLAUDE_VIEW_TYPE = "docx-claude-view";

export class DocxPreviewView extends FileView {
  private currentBuffer: ArrayBuffer | null = null;
  lastLocator: StructuralLocator | null = null;
  private renderEl: HTMLElement | null = null;
  private driftWarningEl: HTMLElement | null = null;

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
    this.allowNoFile = false;
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

  async onLoadFile(file: TFile): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass("docx-claude-view");

    this.driftWarningEl = this.contentEl.createDiv({
      cls: "docx-claude-drift-warning",
    });
    this.driftWarningEl.style.display = "none";

    this.renderEl = this.contentEl.createDiv({ cls: "docx-claude-render" });

    await this.renderFile(file);
    this.attachSelectionListeners();
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    this.detachSelectionListeners();
    this.contentEl.empty();
    this.currentBuffer = null;
    this.lastLocator = null;
    this.renderEl = null;
    this.driftWarningEl = null;
  }

  private async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.renderEl.empty();
    const buf = await this.app.vault.readBinary(file);
    this.currentBuffer = buf;
    await renderAsync(buf, this.renderEl, this.renderEl, {
      className: "docx-claude",
      ignoreLastRenderedPageBreak: true,
      experimental: true,
    });
  }

  async refresh(): Promise<void> {
    if (this.file) {
      await this.renderFile(this.file);
    }
  }

  getBuffer(): ArrayBuffer | null {
    return this.currentBuffer;
  }

  showDriftWarning(message: string): void {
    if (!this.driftWarningEl) return;
    this.driftWarningEl.setText(message);
    this.driftWarningEl.style.display = "block";
  }

  clearDriftWarning(): void {
    if (!this.driftWarningEl) return;
    this.driftWarningEl.style.display = "none";
    this.driftWarningEl.setText("");
  }

  private selectionHandler = (): void => {
    const sel = this.contentEl.win.getSelection();
    if (!sel || sel.rangeCount === 0 || !this.renderEl) {
      this.lastLocator = null;
      return;
    }
    const range = sel.getRangeAt(0);
    if (!this.renderEl.contains(range.startContainer)) return;
    if (range.collapsed) {
      this.lastLocator = null;
      return;
    }
    this.lastLocator = computeLocator(range, this.renderEl);
  };

  private attachSelectionListeners(): void {
    this.contentEl.addEventListener("mouseup", this.selectionHandler);
    this.contentEl.addEventListener("keyup", this.selectionHandler);
    this.contentEl.ownerDocument.addEventListener(
      "selectionchange",
      this.selectionHandler,
    );
  }

  private detachSelectionListeners(): void {
    this.contentEl.removeEventListener("mouseup", this.selectionHandler);
    this.contentEl.removeEventListener("keyup", this.selectionHandler);
    this.contentEl.ownerDocument.removeEventListener(
      "selectionchange",
      this.selectionHandler,
    );
  }
}
