import { FileView, TFile, WorkspaceLeaf } from "obsidian";
import { renderAsync } from "docx-preview";
import { DEFAULT_ZOOM, MAX_ZOOM, MIN_ZOOM, ZOOM_STEP, clampZoom } from "./settings";

export const DOCX_CLAUDE_VIEW_TYPE = "docx-claude-view";

export class DocxPreviewView extends FileView {
  private renderEl: HTMLElement | null = null;
  private zoomIndicatorEl: HTMLElement | null = null;
  private zoom = DEFAULT_ZOOM;

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

  setInitialZoom(z: number): void {
    this.zoom = clampZoom(z);
  }

  async onLoadFile(file: TFile): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass("docx-claude-view");

    const toolbar = this.contentEl.createDiv({ cls: "docx-claude-toolbar" });
    this.buildToolbar(toolbar);

    this.renderEl = this.contentEl.createDiv({ cls: "docx-claude-render" });
    this.applyZoom();

    this.registerDomEvent(this.contentEl, "wheel", this.onWheel, {
      passive: false,
    });

    await this.renderFile(file);
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    this.contentEl.empty();
    this.renderEl = null;
    this.zoomIndicatorEl = null;
  }

  private async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.renderEl.empty();
    const buf = await this.app.vault.readBinary(file);
    await renderAsync(buf, this.renderEl, this.renderEl, {
      className: "docx-claude",
      ignoreLastRenderedPageBreak: true,
      experimental: true,
    });
  }

  private buildToolbar(toolbar: HTMLElement): void {
    const mk = (text: string, title: string, fn: () => void) => {
      const b = toolbar.createEl("button", {
        text,
        cls: "docx-claude-zoom-btn",
        attr: { title, "aria-label": title },
      });
      b.addEventListener("click", (ev) => {
        ev.preventDefault();
        fn();
      });
      return b;
    };
    mk("−", "Zoom out", () => this.zoomOut());
    this.zoomIndicatorEl = toolbar.createDiv({ cls: "docx-claude-zoom-pct" });
    mk("+", "Zoom in", () => this.zoomIn());
    mk("100%", "Reset zoom", () => this.resetZoom());
    this.updateZoomIndicator();
  }

  zoomIn(): void {
    this.setZoom(this.zoom + ZOOM_STEP);
  }

  zoomOut(): void {
    this.setZoom(this.zoom - ZOOM_STEP);
  }

  resetZoom(): void {
    this.setZoom(DEFAULT_ZOOM);
  }

  setZoom(z: number): void {
    const next = clampZoom(Math.round(z * 100) / 100);
    if (next === this.zoom) return;
    this.zoom = next;
    this.applyZoom();
    this.updateZoomIndicator();
  }

  getZoom(): number {
    return this.zoom;
  }

  private applyZoom(): void {
    if (!this.renderEl) return;
    // CSS `zoom` works in Electron/Chromium and scales both visual and layout,
    // so scrollbars and click targets stay correct. Fall back to transform
    // would require width compensation; keep this simple.
    (this.renderEl.style as CSSStyleDeclaration & { zoom?: string }).zoom =
      String(this.zoom);
  }

  private updateZoomIndicator(): void {
    if (!this.zoomIndicatorEl) return;
    this.zoomIndicatorEl.setText(`${Math.round(this.zoom * 100)}%`);
    const atMin = this.zoom <= MIN_ZOOM + 1e-9;
    const atMax = this.zoom >= MAX_ZOOM - 1e-9;
    this.zoomIndicatorEl.toggleClass("at-bound", atMin || atMax);
  }

  private onWheel = (ev: WheelEvent): void => {
    // Cmd/Ctrl + wheel = zoom. Otherwise let the page scroll normally.
    if (!(ev.metaKey || ev.ctrlKey)) return;
    ev.preventDefault();
    if (ev.deltaY < 0) this.zoomIn();
    else if (ev.deltaY > 0) this.zoomOut();
  };
}
