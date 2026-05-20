import { TFile } from "obsidian";
import * as XLSX from "xlsx";
import { OfficeFileView } from "./OfficeFileView";

export const XLSX_CLAUDE_VIEW_TYPE = "xlsx-claude-view";

export class XlsxPreviewView extends OfficeFileView {
  private workbook: XLSX.WorkBook | null = null;
  private sheetEl: HTMLElement | null = null;
  private tabsEl: HTMLElement | null = null;

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
    this.workbook = null;
    this.sheetEl = null;
    this.tabsEl = null;
    await super.onUnloadFile(file);
  }

  protected async renderFile(file: TFile): Promise<void> {
    if (!this.renderEl) return;
    this.renderEl.empty();
    const buf = await this.app.vault.readBinary(file);
    this.workbook = XLSX.read(new Uint8Array(buf), { type: "array" });

    this.sheetEl = this.renderEl.createDiv({ cls: "docx-claude-xlsx-sheet" });

    const names = this.workbook.SheetNames;
    if (names.length > 1) {
      this.tabsEl = createDiv({ cls: "docx-claude-xlsx-tabs" });
      this.buildTabs(names);
      this.appendBelowRender(this.tabsEl);
    }

    if (names.length > 0) this.showSheet(names[0]);
  }

  private buildTabs(names: string[]): void {
    if (!this.tabsEl) return;
    this.tabsEl.empty();
    for (const name of names) {
      const tab = this.tabsEl.createEl("button", {
        text: name,
        cls: "docx-claude-xlsx-tab",
        attr: { title: name, "aria-label": `Switch to sheet ${name}` },
      });
      tab.dataset.sheet = name;
      tab.addEventListener("click", (ev) => {
        ev.preventDefault();
        this.showSheet(name);
      });
    }
  }

  private showSheet(name: string): void {
    if (!this.workbook || !this.sheetEl) return;
    const ws = this.workbook.Sheets[name];
    if (!ws) return;
    this.sheetEl.innerHTML = XLSX.utils.sheet_to_html(ws);
    this.updateActiveTab(name);
  }

  private updateActiveTab(active: string): void {
    if (!this.tabsEl) return;
    const tabs = this.tabsEl.querySelectorAll<HTMLElement>(".docx-claude-xlsx-tab");
    tabs.forEach((t) => t.toggleClass("is-active", t.dataset.sheet === active));
  }
}
