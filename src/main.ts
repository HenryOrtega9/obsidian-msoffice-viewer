import { Plugin, WorkspaceLeaf } from "obsidian";
import { OfficeFileView } from "./OfficeFileView";
import { DOCX_CLAUDE_VIEW_TYPE, DocxPreviewView } from "./DocxPreviewView";
import { PPTX_CLAUDE_VIEW_TYPE, PptxPreviewView } from "./PptxPreviewView";
import { XLSX_CLAUDE_VIEW_TYPE, XlsxPreviewView } from "./XlsxPreviewView";
import {
  DEFAULT_SETTINGS,
  DocxPreviewSettingTab,
  DocxPreviewSettings,
} from "./settings";

export default class MsOfficeViewerPlugin extends Plugin {
  settings: DocxPreviewSettings = DEFAULT_SETTINGS;

  async onload(): Promise<void> {
    await this.loadSettings();

    this.registerView(DOCX_CLAUDE_VIEW_TYPE, (leaf: WorkspaceLeaf) => {
      const view = new DocxPreviewView(leaf);
      view.setInitialZoom(this.settings.defaultZoom);
      return view;
    });
    this.registerExtensions(["docx"], DOCX_CLAUDE_VIEW_TYPE);

    this.registerView(PPTX_CLAUDE_VIEW_TYPE, (leaf: WorkspaceLeaf) => {
      const view = new PptxPreviewView(leaf);
      view.setInitialZoom(this.settings.defaultZoom);
      return view;
    });
    this.registerExtensions(["pptx"], PPTX_CLAUDE_VIEW_TYPE);

    this.registerView(XLSX_CLAUDE_VIEW_TYPE, (leaf: WorkspaceLeaf) => {
      const view = new XlsxPreviewView(leaf);
      view.setInitialZoom(this.settings.defaultZoom);
      return view;
    });
    this.registerExtensions(["xlsx", "xls"], XLSX_CLAUDE_VIEW_TYPE);

    this.addCommand({
      id: "docx-zoom-in",
      name: "Zoom in",
      checkCallback: (checking) => this.withActiveView(checking, (v) => v.zoomIn()),
    });

    this.addCommand({
      id: "docx-zoom-out",
      name: "Zoom out",
      checkCallback: (checking) => this.withActiveView(checking, (v) => v.zoomOut()),
    });

    this.addCommand({
      id: "docx-zoom-reset",
      name: "Reset zoom",
      checkCallback: (checking) => this.withActiveView(checking, (v) => v.resetZoom()),
    });

    this.addSettingTab(new DocxPreviewSettingTab(this.app, this));
  }

  onunload(): void {
    // Obsidian unregisters view/commands automatically.
  }

  async loadSettings(): Promise<void> {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  private withActiveView(
    checking: boolean,
    action: (v: OfficeFileView) => void,
  ): boolean {
    const view =
      this.app.workspace.getActiveViewOfType(DocxPreviewView) ??
      this.app.workspace.getActiveViewOfType(PptxPreviewView) ??
      this.app.workspace.getActiveViewOfType(XlsxPreviewView);
    if (!view || !view.file) return false;
    if (!checking) action(view);
    return true;
  }
}
