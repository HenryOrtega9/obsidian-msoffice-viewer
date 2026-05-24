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
import type { InternalLinkTarget } from "./xlsx/hyperlinks";

export const XLSX_CLAUDE_VIEW_TYPE = "xlsx-claude-view";

interface SheetEntry {
  name: string;
  worksheet: ExcelJS.Worksheet;
  hidden: boolean;
}

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

    if (file.extension === "xlsx") {
      try {
        await this.renderViaExcelJsGrid(file);
        return;
      } catch (e) {
        console.error("ExcelJS grid render failed; falling back to LibreOffice PDF:", e);
        new Notice("Grid rendering failed; using PDF fallback.");
        if (this.file !== file || !this.renderEl) return;
        // resetState before empty() so a half-built tabsEl (which lives
        // outside renderEl) doesn't ghost beneath the PDF fallback.
        this.resetState();
        this.renderEl.empty();
      }
    }

    // .xls (legacy binary) or ExcelJS fallback path
    const sofficeBin = await findSoffice();
    if (sofficeBin) {
      try {
        await this.renderViaLibreOfficePdf(file, sofficeBin, file.extension);
        return;
      } catch (e) {
        console.error("LibreOffice render failed:", e);
      }
    }

    if (this.renderEl) {
      this.renderEl
        .createDiv({ cls: "docx-claude-pdf-error" })
        .setText(
          file.extension === "xls"
            ? "Couldn't render this .xls. Install LibreOffice or convert to .xlsx."
            : "Couldn't render this workbook. See console for details.",
        );
    }
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
      this.renderEl
        .createDiv({ cls: "docx-claude-pdf-error" })
        .setText("This workbook has no visible sheets.");
      return;
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
