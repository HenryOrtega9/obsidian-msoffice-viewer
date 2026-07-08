import { Plugin, TFile, WorkspaceLeaf } from "obsidian";
import { OfficeFileView } from "./OfficeFileView";
import { DOCX_CLAUDE_VIEW_TYPE, DocxPreviewView } from "./DocxPreviewView";
import { PPTX_CLAUDE_VIEW_TYPE, PptxPreviewView } from "./PptxPreviewView";
import { XLSX_CLAUDE_VIEW_TYPE, XlsxPreviewView } from "./XlsxPreviewView";
import { ClaudeBridge } from "./ClaudeBridge";
import { CreateModal } from "./CreateModal";
import {
  DEFAULT_SETTINGS,
  DocxPreviewSettingTab,
  DocxPreviewSettings,
  clampZoom,
} from "./settings";
import { patchExcelJsDefaultNumFmts } from "./xlsx/numFmtPatch";

export default class MsOfficeViewerPlugin extends Plugin {
  settings: DocxPreviewSettings = DEFAULT_SETTINGS;

  async onload(): Promise<void> {
    patchExcelJsDefaultNumFmts();
    await this.loadSettings();

    OfficeFileView.pluginId = this.manifest.id;

    this.registerView(DOCX_CLAUDE_VIEW_TYPE, (leaf: WorkspaceLeaf) => {
      const view = new DocxPreviewView(leaf);
      view.setInitialZoom(this.settings.defaultZoom);
      return view;
    });
    this.registerView(PPTX_CLAUDE_VIEW_TYPE, (leaf: WorkspaceLeaf) => {
      const view = new PptxPreviewView(leaf);
      view.setInitialZoom(this.settings.defaultZoom);
      return view;
    });
    this.registerView(XLSX_CLAUDE_VIEW_TYPE, (leaf: WorkspaceLeaf) => {
      const view = new XlsxPreviewView(leaf);
      view.setInitialZoom(this.settings.defaultZoom);
      return view;
    });

    // Each registerExtensions in its own try/catch so a single extension
    // collision (another plugin already owns .xls, say) doesn't abort the
    // rest of onload.
    this.safeRegisterExtensions(["docx"], DOCX_CLAUDE_VIEW_TYPE);
    this.safeRegisterExtensions(["pptx"], PPTX_CLAUDE_VIEW_TYPE);
    this.safeRegisterExtensions(["xlsx"], XLSX_CLAUDE_VIEW_TYPE);
    this.safeRegisterExtensions(["xls"], XLSX_CLAUDE_VIEW_TYPE);

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

    this.addCommand({
      id: "msoffice-create-with-claude",
      name: "Create Office file with Claude",
      callback: () => {
        const bridge = new ClaudeBridge(this.app);
        new CreateModal(this.app, bridge, (file) => this.openOfficeFile(file)).open();
      },
    });

    this.addSettingTab(new DocxPreviewSettingTab(this.app, this));
  }

  // No onunload leaf-detach: Obsidian's plugin guidelines forbid it. Detaching
  // here permanently closes the user's open tabs on every plugin update;
  // Obsidian re-adopts orphaned view types itself when the plugin reloads.

  private safeRegisterExtensions(exts: string[], viewType: string): void {
    try {
      this.registerExtensions(exts, viewType);
    } catch (e) {
      console.warn(
        `Couldn't register extensions ${exts.join(",")} for ${viewType}:`,
        e,
      );
    }
  }

  async loadSettings(): Promise<void> {
    // Whitelist known keys and coerce types so a corrupt/hand-edited
    // data.json can't leak garbage into the runtime settings.
    const raw = (await this.loadData()) as Partial<DocxPreviewSettings> | null;
    const rawZoom = raw && typeof raw === "object" ? raw.defaultZoom : undefined;
    const zoom = clampZoom(Number(rawZoom));
    this.settings = { ...DEFAULT_SETTINGS, defaultZoom: zoom };
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  private async openOfficeFile(file: TFile): Promise<void> {
    const viewType =
      file.extension === "docx"
        ? DOCX_CLAUDE_VIEW_TYPE
        : file.extension === "pptx"
          ? PPTX_CLAUDE_VIEW_TYPE
          : file.extension === "xlsx" || file.extension === "xls"
            ? XLSX_CLAUDE_VIEW_TYPE
            : null;
    if (!viewType) return;
    const leaf = this.app.workspace.getLeaf(true);
    await leaf.setViewState({ type: viewType, state: { file: file.path }, active: true });
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
