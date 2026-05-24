import { FileSystemAdapter, FileView, Notice, TFile, WorkspaceLeaf } from "obsidian";
import { DEFAULT_ZOOM, MAX_ZOOM, MIN_ZOOM, ZOOM_STEP, clampZoom } from "./settings";
import {
  convertOfficeToPdf,
  ensurePdfjsWorker,
  pluginCacheDir,
  renderPdfPagesIntoStage,
  type PdfRenderHandle,
} from "./officeToPdf";

interface ElectronShellLike {
  openPath(path: string): Promise<string>;
}

function loadElectronShell(): ElectronShellLike | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const electron = (window as unknown as { require?: NodeRequire }).require?.("electron");
    return electron?.shell ?? null;
  } catch {
    return null;
  }
}

export abstract class OfficeFileView extends FileView {
  // Set by main.ts at plugin load so the cache dir tracks the manifest id.
  static pluginId = "obsidian-msoffice-viewer";

  protected renderEl: HTMLElement | null = null;
  protected toolbarEl: HTMLElement | null = null;
  private zoomIndicatorEl: HTMLElement | null = null;
  private zoom = DEFAULT_ZOOM;
  private activePdfRender: PdfRenderHandle | null = null;

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
    this.allowNoFile = false;
  }

  protected abstract renderFile(file: TFile): Promise<void>;
  protected abstract getExternalAppLabel(): string;

  setInitialZoom(z: number): void {
    this.zoom = clampZoom(z);
  }

  async onLoadFile(file: TFile): Promise<void> {
    this.cancelActivePdfRender();
    this.contentEl.empty();
    this.contentEl.addClass("docx-claude-view");

    this.toolbarEl = this.contentEl.createDiv({ cls: "docx-claude-toolbar" });
    this.buildToolbar(this.toolbarEl);

    this.renderEl = this.contentEl.createDiv({ cls: "docx-claude-render" });
    this.applyZoom();

    this.registerDomEvent(this.contentEl, "wheel", this.onWheel, {
      passive: false,
    });

    await this.renderFile(file);
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    this.cancelActivePdfRender();
    this.contentEl.empty();
    this.renderEl = null;
    this.toolbarEl = null;
    this.zoomIndicatorEl = null;
  }

  protected cancelActivePdfRender(): void {
    if (this.activePdfRender) {
      try { this.activePdfRender.cancel(); } catch { /* ignore */ }
      this.activePdfRender = null;
    }
  }

  protected appendBelowRender(el: HTMLElement): void {
    this.contentEl.appendChild(el);
  }

  // Shared LibreOffice → PDF → PDF.js pipeline. Throws if every page failed
  // (callers fall back to the JS renderer). The guard against this.file
  // changing lets us bail cleanly if the user switches files mid-render.
  protected async renderViaLibreOfficePdf(
    file: TFile,
    sofficeBin: string,
    ext: string,
  ): Promise<void> {
    if (!this.renderEl) return;
    const loadingStage = this.renderEl.createDiv({ cls: "docx-claude-pdf-stage" });
    loadingStage
      .createDiv({ cls: "docx-claude-pdf-loading" })
      .setText("Rendering with LibreOffice…");

    await ensurePdfjsWorker(this.app, OfficeFileView.pluginId);

    const buf = await this.app.vault.readBinary(file);
    const pdfPath = await convertOfficeToPdf(
      sofficeBin,
      buf,
      ext,
      pluginCacheDir(OfficeFileView.pluginId),
    );

    if (this.file !== file || !this.renderEl) return;
    this.renderEl.empty();
    const stage = this.renderEl.createDiv({ cls: "docx-claude-pdf-stage" });
    this.cancelActivePdfRender();
    const handle = renderPdfPagesIntoStage(
      pdfPath,
      stage,
      "docx-claude-pdf-slide",
      "docx-claude-pdf-canvas",
      { isStale: () => this.file !== file },
    );
    this.activePdfRender = handle;
    try {
      const pages = await handle.pages;
      if (this.activePdfRender === handle) this.activePdfRender = null;
      if (pages.length === 0 || pages.every((p) => p.failed)) {
        throw new Error("LibreOffice PDF produced no renderable pages.");
      }
    } catch (e) {
      if (this.activePdfRender === handle) this.activePdfRender = null;
      throw e;
    }
  }

  private buildToolbar(toolbar: HTMLElement): void {
    const mkZoom = (text: string, title: string, fn: () => void) => {
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
    mkZoom("−", "Zoom out", () => this.zoomOut());
    this.zoomIndicatorEl = toolbar.createDiv({ cls: "docx-claude-zoom-pct" });
    mkZoom("+", "Zoom in", () => this.zoomIn());
    mkZoom("100%", "Reset zoom", () => this.resetZoom());
    this.updateZoomIndicator();

    this.buildExtraToolbar(toolbar);

    if (loadElectronShell() && this.app.vault.adapter instanceof FileSystemAdapter) {
      const label = this.getExternalAppLabel();
      const openBtn = toolbar.createEl("button", {
        text: label,
        cls: "docx-claude-zoom-btn docx-claude-open-external",
        attr: { title: label, "aria-label": label },
      });
      openBtn.addEventListener("click", (ev) => {
        ev.preventDefault();
        void this.openExternal();
      });
    }
  }

  // Hook for subclasses to add their own toolbar controls (e.g. the docx
  // high-fidelity toggle). Called after the zoom buttons, before "Open in app".
  protected buildExtraToolbar(_toolbar: HTMLElement): void {}

  private async openExternal(): Promise<void> {
    if (!this.file) return;
    const shell = loadElectronShell();
    if (!shell) {
      new Notice("Opening external apps isn't available on this platform.");
      return;
    }
    const adapter = this.app.vault.adapter;
    if (!(adapter instanceof FileSystemAdapter)) {
      new Notice("Vault is not on the local filesystem.");
      return;
    }
    const absPath = adapter.getFullPath(this.file.path);
    const err = await shell.openPath(absPath);
    if (err) new Notice(err);
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
    (this.renderEl.style as CSSStyleDeclaration & { zoom?: string }).zoom = String(this.zoom);
  }

  private updateZoomIndicator(): void {
    if (!this.zoomIndicatorEl) return;
    this.zoomIndicatorEl.setText(`${Math.round(this.zoom * 100)}%`);
    const atMin = this.zoom <= MIN_ZOOM + 1e-9;
    const atMax = this.zoom >= MAX_ZOOM - 1e-9;
    this.zoomIndicatorEl.toggleClass("at-bound", atMin || atMax);
  }

  private onWheel = (ev: WheelEvent): void => {
    if (!(ev.metaKey || ev.ctrlKey)) return;
    ev.preventDefault();
    if (ev.deltaY < 0) this.zoomIn();
    else if (ev.deltaY > 0) this.zoomOut();
  };
}
