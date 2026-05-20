import { Plugin, WorkspaceLeaf } from "obsidian";
import { DOCX_CLAUDE_VIEW_TYPE, DocxPreviewView } from "./DocxPreviewView";
import {
  DEFAULT_SETTINGS,
  DocxPreviewSettingTab,
  DocxPreviewSettings,
} from "./settings";

export default class DocxClaudePlugin extends Plugin {
  settings: DocxPreviewSettings = DEFAULT_SETTINGS;

  async onload(): Promise<void> {
    await this.loadSettings();

    this.registerView(DOCX_CLAUDE_VIEW_TYPE, (leaf: WorkspaceLeaf) => {
      const view = new DocxPreviewView(leaf);
      view.setInitialZoom(this.settings.defaultZoom);
      return view;
    });
    this.registerExtensions(["docx"], DOCX_CLAUDE_VIEW_TYPE);

    this.addCommand({
      id: "docx-zoom-in",
      name: "Docx: Zoom in",
      checkCallback: (checking) => this.withActiveView(checking, (v) => v.zoomIn()),
    });

    this.addCommand({
      id: "docx-zoom-out",
      name: "Docx: Zoom out",
      checkCallback: (checking) => this.withActiveView(checking, (v) => v.zoomOut()),
    });

    this.addCommand({
      id: "docx-zoom-reset",
      name: "Docx: Reset zoom",
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
    action: (v: DocxPreviewView) => void,
  ): boolean {
    const view = this.app.workspace.getActiveViewOfType(DocxPreviewView);
    if (!view || !view.file) return false;
    if (!checking) action(view);
    return true;
  }
}
