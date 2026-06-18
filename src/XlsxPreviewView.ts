import { Notice, TFile } from "obsidian";
import ExcelJS from "exceljs";
import { OfficeFileView } from "./OfficeFileView";
import { findSoffice } from "./officeToPdf";
import { renderSheetIntoGrid, type GridContext } from "./xlsx/grid";
import { applyFrozenPanes } from "./xlsx/panes";
import { renderSheetImages } from "./xlsx/images";
import { applyConditionalFormatting } from "./xlsx/conditionalFormatting";
import { loadWorkbookCharts, type ChartPlacement } from "./xlsx/charts";
import { renderSheetCharts } from "./xlsx/charts/render";
import type { Chart } from "chart.js";
import { ExcelColorRef, resolveExcelColor } from "./xlsx/colors";
import { loadWorkbookTheme } from "./xlsx/themes";
import { colNum } from "./xlsx/merges";
import { closeActivePopover } from "./xlsx/notes";
import type { InternalLinkTarget } from "./xlsx/hyperlinks";

export const XLSX_CLAUDE_VIEW_TYPE = "xlsx-claude-view";

interface SheetEntry {
  name: string;
  worksheet: ExcelJS.Worksheet;
  hidden: boolean;
}

type RenderMode = "pdf" | "grid";

export class XlsxPreviewView extends OfficeFileView {
  private workbook: ExcelJS.Workbook | null = null;
  private sheets: SheetEntry[] = [];
  private gridEl: HTMLElement | null = null;
  private tabsEl: HTMLElement | null = null;
  private gridContext: GridContext | null = null;
  private showHidden = false;
  private activeSheet: string | null = null;
  private theme: string[] | null = null;
  private imageObjectUrls: string[] = [];
  private chartMap: Map<string, ChartPlacement[]> = new Map();
  private chartInstances: Chart[] = [];
  // PDF (high fidelity) vs the interactive ExcelJS grid (sheet tabs, links,
  // selection). userForcedMode is a sticky manual override (null = auto-route),
  // reset on file switch and preserved across toggle-driven re-renders.
  private renderMode: RenderMode = "pdf";
  private userForcedMode: RenderMode | null = null;
  private toggleBtn: HTMLButtonElement | null = null;
  private sofficeAvailable = false;

  async onLoadFile(file: TFile): Promise<void> {
    this.userForcedMode = null;
    await super.onLoadFile(file);
  }

  getViewType(): string {
    return XLSX_CLAUDE_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Xlsx preview";
  }

  getIcon(): string {
    return "table";
  }

  canAcceptExtension(extension: string): boolean {
    return extension === "xlsx" || extension === "xls";
  }

  protected getExternalAppLabel(): string {
    return "Open in Excel";
  }

  async onUnloadFile(file: TFile): Promise<void> {
    this.resetState();
    await super.onUnloadFile(file);
  }

  private resetState(): void {
    // Tear down any in-flight PDF render first: resetState runs on every
    // renderFile entry (file switch, PDF→grid toggle, PDF-fail→grid fallback),
    // so without this the PDF.js document + IntersectionObserver leak whenever
    // we leave PDF mode for the grid. Idempotent, so the extra calls already in
    // renderViaLibreOfficePdf stay harmless no-ops.
    this.cancelActivePdfRender();
    closeActivePopover();
    this.revokeImageUrls();
    this.destroyCharts();
    this.workbook = null;
    this.sheets = [];
    this.gridEl = null;
    this.gridContext = null;
    this.showHidden = false;
    this.activeSheet = null;
    this.theme = null;
    this.chartMap = new Map();
    if (this.tabsEl) {
      this.tabsEl.remove();
      this.tabsEl = null;
    }
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

  protected async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.resetState();
    this.renderEl.empty();

    // LibreOffice → PDF is the high-fidelity default for every spreadsheet type
    // (it renders through the real Excel-compatible layout engine). The ExcelJS
    // grid (sheet tabs, hyperlinks, text selection) is the fallback when
    // LibreOffice is unavailable or fails, and is reachable via the toolbar
    // toggle. The grid only handles .xlsx, not the legacy .xls binary format.
    const canGrid = file.extension === "xlsx";
    const mode: RenderMode = this.userForcedMode ?? "pdf";

    if (mode === "grid" && canGrid) {
      if (await this.tryRenderGrid(file)) return;
      if (this.file !== file || !this.renderEl) return;
      this.resetState();
      this.renderEl.empty();
      if (await this.tryRenderPdf(file)) return;
      this.showError(file);
      return;
    }

    if (await this.tryRenderPdf(file)) return;
    if (this.file !== file || !this.renderEl) return;
    // resetState before empty() so a half-built tabsEl (which lives outside
    // renderEl) doesn't ghost beneath the fallback.
    this.resetState();
    this.renderEl.empty();
    if (canGrid && (await this.tryRenderGrid(file))) return;
    this.showError(file);
  }

  // Render via the shared LibreOffice -> PDF -> PDF.js pipeline. Returns false
  // when LibreOffice is missing or the conversion fails.
  private async tryRenderPdf(file: TFile): Promise<boolean> {
    if (!this.renderEl) return false;
    const sofficeBin = await findSoffice();
    if (this.file !== file || !this.renderEl) return false;
    this.sofficeAvailable = sofficeBin != null;
    if (!sofficeBin) return false;
    this.resetState();
    this.renderEl.empty();
    try {
      await this.renderViaLibreOfficePdf(file, sofficeBin, file.extension);
      // Re-check freshness: a rapid file switch during the await would otherwise
      // write renderMode/toggle label for the wrong file. The stale renderFile
      // caller then hits its own this.file guard and stops.
      if (this.file !== file || !this.renderEl) return false;
      this.renderMode = "pdf";
      this.updateToggleLabel();
      return true;
    } catch (e) {
      console.error("LibreOffice render failed:", e);
      return false;
    }
  }

  // Render via the interactive ExcelJS grid. Returns false on parse/render
  // failure so the caller can fall back to PDF.
  private async tryRenderGrid(file: TFile): Promise<boolean> {
    if (!this.renderEl) return false;
    try {
      await this.renderViaExcelJsGrid(file);
      // Re-check freshness after the async parse so a rapid file switch can't
      // write renderMode/toggle label for the wrong file.
      if (this.file !== file || !this.renderEl) return false;
      this.renderMode = "grid";
      this.updateToggleLabel();
      return true;
    } catch (e) {
      console.error("ExcelJS grid render failed:", e);
      return false;
    }
  }

  private showError(file: TFile): void {
    if (!this.renderEl) return;
    this.renderEl.empty();
    this.renderEl
      .createDiv({ cls: "docx-claude-pdf-error" })
      .setText(
        file.extension === "xls"
          ? "Couldn't render this .xls. Install LibreOffice or convert to .xlsx."
          : "Couldn't render this workbook. See console for details.",
      );
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
    this.updateToggleLabel();
  }

  private toggleLabel(): string {
    return this.renderMode === "pdf" ? "Interactive grid" : "High fidelity (PDF)";
  }

  private toggleTitle(): string {
    return this.renderMode === "pdf"
      ? "Switch to the interactive grid (sheet tabs, hyperlinks, text selection)"
      : "Re-render this workbook via LibreOffice for higher fidelity";
  }

  private updateToggleLabel(): void {
    if (!this.toggleBtn) return;
    // The grid handles .xlsx only; for .xls there is no alternative to PDF, so
    // hide the toggle entirely.
    this.toggleBtn.style.display = this.file?.extension === "xlsx" ? "" : "none";
    this.toggleBtn.setText(this.toggleLabel());
    this.toggleBtn.setAttribute("title", this.toggleTitle());
    this.toggleBtn.setAttribute("aria-label", this.toggleTitle());
  }

  private async onToggleClick(): Promise<void> {
    if (!this.file) return;
    const target: RenderMode = this.renderMode === "pdf" ? "grid" : "pdf";
    if (target === "grid" && this.file.extension !== "xlsx") {
      new Notice("The interactive grid supports .xlsx only. Use Open in Excel for .xls.", 6000);
      return;
    }
    if (target === "pdf") {
      const soffice = await findSoffice();
      if (!soffice) {
        new Notice("LibreOffice not found. Install it for high-fidelity PDF rendering.", 6000);
        return;
      }
    }
    this.userForcedMode = target;
    await this.renderFile(this.file);
  }

  private async renderViaExcelJsGrid(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    const buf = await this.app.vault.readBinary(file);
    const wb = new ExcelJS.Workbook();
    const [, theme] = await Promise.all([
      wb.xlsx.load(buf),
      loadWorkbookTheme(buf),
    ]);
    if (this.file !== file || !this.renderEl) return;
    this.chartMap = await loadWorkbookCharts(buf, theme ?? undefined);
    if (this.file !== file || !this.renderEl) return;
    this.workbook = wb;
    this.theme = theme;

    this.sheets = [];
    wb.worksheets.forEach((ws) => {
      // veryHidden sheets are excluded entirely (user can't unhide them in
      // Excel without code). Hidden sheets are stashed but surface-able via
      // the "Show hidden" toggle.
      if (ws.state === "veryHidden") return;
      this.sheets.push({
        name: ws.name,
        worksheet: ws,
        hidden: ws.state === "hidden",
      });
    });

    const visibleCount = this.sheets.filter((s) => !s.hidden).length;
    if (visibleCount === 0 && this.sheets.length === 0) {
      // Throw rather than render an error div: tryRenderGrid then returns false
      // so the caller falls back to the PDF path (which may still render this
      // workbook), instead of the grid masking a renderable file with an error.
      throw new Error("xlsx-grid: no renderable sheets");
    }
    if (visibleCount === 0) {
      // All sheets are hidden — auto-enable showHidden so the user sees something.
      this.showHidden = true;
    }

    this.gridEl = this.renderEl.createDiv({ cls: "docx-claude-xlsx-grid" });

    if (this.sheets.length > 1 || this.sheets.some((s) => s.hidden)) {
      this.tabsEl = createDiv({ cls: "docx-claude-xlsx-tabs" });
      this.buildTabs();
      this.appendBelowRender(this.tabsEl);
    }

    const initial = this.sheets.find((s) => !s.hidden) ?? this.sheets[0];
    if (initial) this.showSheet(initial.name);
  }

  private buildTabs(): void {
    if (!this.tabsEl) return;
    this.tabsEl.empty();
    const anyHidden = this.sheets.some((s) => s.hidden);

    for (const entry of this.sheets) {
      if (entry.hidden && !this.showHidden) continue;
      const tab = this.tabsEl.createEl("button", {
        text: entry.name,
        cls:
          "docx-claude-xlsx-tab" + (entry.hidden ? " is-hidden-sheet" : ""),
        attr: {
          title: entry.hidden ? `${entry.name} (hidden)` : entry.name,
          "aria-label": `Switch to sheet ${entry.name}`,
        },
      });
      tab.dataset.sheet = entry.name;
      const tabColorRef = (entry.worksheet as unknown as {
        properties?: { tabColor?: ExcelColorRef };
      }).properties?.tabColor;
      const tabColor = resolveExcelColor(tabColorRef, this.theme ?? undefined);
      if (tabColor) tab.style.borderBottom = `3px solid ${tabColor}`;
      tab.addEventListener("click", (ev) => {
        ev.preventDefault();
        this.showSheet(entry.name);
      });
    }

    if (anyHidden) {
      const toggle = this.tabsEl.createEl("button", {
        text: this.showHidden ? "Hide hidden sheets" : "Show hidden sheets",
        cls: "docx-claude-xlsx-tab docx-claude-xlsx-toggle-hidden",
        attr: {
          title: "Toggle visibility of sheets marked hidden in Excel",
          "aria-label": "Toggle hidden sheets",
        },
      });
      toggle.addEventListener("click", (ev) => {
        ev.preventDefault();
        this.showHidden = !this.showHidden;
        this.buildTabs();
        if (this.activeSheet) this.updateActiveTab(this.activeSheet);
      });
    }
  }

  private showSheet(name: string): void {
    if (!this.gridEl) return;
    const entry = this.sheets.find((s) => s.name === name);
    if (!entry) return;
    closeActivePopover();
    this.revokeImageUrls();
    this.destroyCharts();
    this.gridEl.empty();
    this.gridContext = renderSheetIntoGrid(entry.worksheet, this.gridEl, {
      onInternalLink: (target) => this.followInternalLink(target),
      theme: this.theme ?? undefined,
    });
    applyConditionalFormatting(entry.worksheet, this.gridContext, this.theme ?? undefined);
    applyFrozenPanes(entry.worksheet, this.gridContext);
    if (this.workbook) {
      const result = renderSheetImages(entry.worksheet, this.workbook, this.gridContext);
      this.imageObjectUrls = result.objectUrls;
    }
    const placements = this.chartMap.get(name);
    if (placements && placements.length > 0) {
      this.chartInstances = renderSheetCharts(placements, this.gridContext, {
        onRequestLibreOffice: () => void this.renderViaLibreOfficeFallback(),
      });
    }
    this.activeSheet = name;
    this.updateActiveTab(name);
  }

  private async renderViaLibreOfficeFallback(): Promise<void> {
    if (!this.file || !this.renderEl) return;
    const file = this.file;
    const soffice = await findSoffice();
    if (!soffice) {
      new Notice("LibreOffice not found. Install it to render this chart faithfully.");
      return;
    }
    if (this.file !== file || !this.renderEl) return;
    closeActivePopover();
    this.destroyCharts();
    this.revokeImageUrls();
    if (this.tabsEl) {
      this.tabsEl.remove();
      this.tabsEl = null;
    }
    this.gridEl = null;
    this.gridContext = null;
    this.renderEl.empty();
    try {
      await this.renderViaLibreOfficePdf(file, soffice, file.extension);
      // The view now shows the PDF, so sync the mode, sticky override, and
      // toggle label (mirrors tryRenderPdf). Without this the toggle still reads
      // "Interactive grid"→grid was the last mode, trapping the view in PDF with
      // a wrong label. Set only on success so a failed convert doesn't lie.
      this.renderMode = "pdf";
      this.userForcedMode = "pdf";
      this.updateToggleLabel();
    } catch (e) {
      console.error("LibreOffice fallback failed:", e);
      new Notice("LibreOffice rendering failed. See console for details.");
    }
  }

  private followInternalLink(target: InternalLinkTarget): void {
    const entry = this.sheets.find((s) => s.name === target.sheet);
    if (!entry) {
      new Notice(`Link target sheet "${target.sheet}" not found.`);
      return;
    }
    // If the target is a hidden sheet, surface it temporarily.
    if (entry.hidden && !this.showHidden) {
      this.showHidden = true;
      this.buildTabs();
    }
    this.showSheet(target.sheet);
    const cellMatch = target.address.match(/^([A-Z]+)(\d+)$/);
    if (cellMatch && this.gridContext) {
      const col = colNum(cellMatch[1]);
      const row = parseInt(cellMatch[2], 10);
      const td = this.gridContext.cellMap.get(`${row}:${col}`);
      if (td) {
        td.scrollIntoView({ block: "center", inline: "center" });
        td.classList.add("docx-claude-xlsx-cell-flash");
        window.setTimeout(() => td.classList.remove("docx-claude-xlsx-cell-flash"), 1200);
      }
    }
  }

  private updateActiveTab(active: string): void {
    if (!this.tabsEl) return;
    const tabs = this.tabsEl.querySelectorAll<HTMLElement>(".docx-claude-xlsx-tab");
    tabs.forEach((t) => t.toggleClass("is-active", t.dataset.sheet === active));
  }
}
