import { Notice, Plugin, TFile, WorkspaceLeaf, normalizePath } from "obsidian";
import { spawn } from "child_process";
import * as mammoth from "mammoth";
import {
  DOCX_CLAUDE_VIEW_TYPE,
  DocxPreviewView,
} from "./DocxPreviewView";
import { ClaudeBridge } from "./ClaudeBridge";
import { BackupManager } from "./BackupManager";
import { dispatch } from "./EditDispatcher";
import { EditInstructionModal } from "./EditModal";
import {
  DEFAULT_SETTINGS,
  DocxClaudeSettings,
  DocxClaudeSettingTab,
} from "./settings";

export default class DocxClaudePlugin extends Plugin {
  settings: DocxClaudeSettings = DEFAULT_SETTINGS;
  bridge!: ClaudeBridge;
  backup!: BackupManager;

  async onload(): Promise<void> {
    await this.loadSettings();

    const pluginDir = this.manifest.dir
      ? normalizePath(this.manifest.dir)
      : normalizePath(
          `${this.app.vault.configDir}/plugins/${this.manifest.id}`,
        );

    this.bridge = new ClaudeBridge(this.app);
    this.backup = new BackupManager(
      this.app,
      pluginDir,
      this.settings.maxHistory,
    );

    this.registerView(
      DOCX_CLAUDE_VIEW_TYPE,
      (leaf: WorkspaceLeaf) => new DocxPreviewView(leaf),
    );
    this.registerExtensions(["docx"], DOCX_CLAUDE_VIEW_TYPE);

    this.addCommand({
      id: "docx-claude-ask-edit",
      name: "Docx: Ask Claude to edit selection",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(DocxPreviewView);
        if (!view || !view.file) return false;
        if (checking) return true;
        this.askClaudeForActiveSelection(view);
        return true;
      },
    });

    this.addCommand({
      id: "docx-claude-convert-md",
      name: "Docx: Convert to markdown note (one-shot)",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(DocxPreviewView);
        if (!view || !view.file) return false;
        if (checking) return true;
        this.convertToMarkdownNote(view.file);
        return true;
      },
    });

    this.addCommand({
      id: "docx-claude-debug-locator",
      name: "Docx: Log current selection locator (debug)",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(DocxPreviewView);
        if (!view || !view.file) return false;
        if (checking) return true;
        // eslint-disable-next-line no-console
        console.log("[docx-claude] locator:", view.lastLocator);
        new Notice(
          view.lastLocator
            ? `Locator: paragraph ${view.lastLocator.paragraphIndex} ` +
                `[${view.lastLocator.startOffset}..${view.lastLocator.endOffset}]`
            : "No selection",
        );
        return true;
      },
    });

    this.addSettingTab(new DocxClaudeSettingTab(this.app, this));

    this.checkPandoc();
  }

  onunload(): void {
    // Obsidian unregisters the view and command on unload automatically.
  }

  async loadSettings(): Promise<void> {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
    this.backup.setMaxHistory(this.settings.maxHistory);
  }

  private async askClaudeForActiveSelection(view: DocxPreviewView): Promise<void> {
    const locator = view.lastLocator;
    const file = view.file;
    if (!locator || !file) {
      new Notice("Select text in the docx preview first.");
      return;
    }
    if (!this.bridge.isAvailable()) {
      new Notice(
        "claude-cli-chat plugin not found or missing a programmatic prompt method. " +
          "Install / enable / upgrade it.",
      );
      return;
    }

    new EditInstructionModal(
      this.app,
      locator,
      async (instruction) => {
        try {
          new Notice("Asking Claude...");
          const response = await this.bridge.askClaude(locator, instruction);
          const outcome = await dispatch(
            {
              app: this.app,
              pluginDir: this.manifest.dir ?? "",
              backup: this.backup,
              settings: this.settings,
            },
            file,
            response,
            locator,
          );
          await view.refresh();
          if (outcome.warning) {
            view.showDriftWarning(outcome.warning);
          } else {
            view.clearDriftWarning();
          }
          new Notice(
            `Edit applied via ${outcome.pathUsed} path (${outcome.bytesWritten} bytes).`,
          );
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          new Notice(`Edit failed: ${message}`);
          // eslint-disable-next-line no-console
          console.error("[docx-claude] edit failed", err);
        }
      },
      () => {
        /* cancelled */
      },
    ).open();
  }

  private async convertToMarkdownNote(file: TFile): Promise<void> {
    try {
      const buf = await this.app.vault.readBinary(file);
      const result = await mammoth.convertToMarkdown({ buffer: Buffer.from(buf) });
      const mdPath = file.path.replace(/\.docx$/i, ".md");
      const existing = this.app.vault.getAbstractFileByPath(mdPath);
      if (existing && existing instanceof TFile) {
        await this.app.vault.modify(existing, result.value);
      } else {
        await this.app.vault.create(mdPath, result.value);
      }
      new Notice(`Wrote ${mdPath}`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      new Notice(`Conversion failed: ${message}`);
    }
  }

  private checkPandoc(): void {
    const child = spawn(this.settings.pandocPath, ["--version"]);
    let resolved = false;
    const fail = (msg: string) => {
      if (resolved) return;
      resolved = true;
      new Notice(
        `obsidian-docx-claude: pandoc not found (${msg}). Round-trip edits will fail. ` +
          `Install pandoc (e.g. brew install pandoc) and set the path in plugin settings.`,
        10000,
      );
    };
    child.on("error", (err) => fail(err.message));
    child.on("close", (code) => {
      if (code !== 0 && !resolved) fail(`exit code ${code}`);
      else resolved = true;
    });
  }
}
