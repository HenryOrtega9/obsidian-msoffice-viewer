import { App, Modal, Notice, Setting, TFile, TFolder, normalizePath } from "obsidian";
import { ClaudeBridge, ClaudeBridgeError } from "./ClaudeBridge";
import {
  buildFile,
  extractJsonFromResponse,
  FileKind,
  SpecValidationError,
  systemPromptFor,
  validateSpec,
} from "./generators";

const ILLEGAL_FILENAME_CHARS = /[\\/:*?"<>|]/g;

interface OpenFn {
  (file: TFile): Promise<void> | void;
}

export class CreateModal extends Modal {
  private kind: FileKind = "docx";
  private filename = "Untitled.docx";
  private filenameTouched = false;
  private description = "";
  private busy = false;
  private statusEl: HTMLElement | null = null;
  private submitBtn: HTMLButtonElement | null = null;
  private filenameInput: HTMLInputElement | null = null;

  constructor(
    app: App,
    private readonly bridge: ClaudeBridge,
    private readonly openFile: OpenFn,
  ) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h2", { text: "Create Office file with Claude" });

    new Setting(contentEl)
      .setName("File type")
      .addDropdown((dd) => {
        dd.addOption("docx", "Word (.docx)")
          .addOption("pptx", "PowerPoint (.pptx)")
          .addOption("xlsx", "Excel (.xlsx)")
          .setValue(this.kind)
          .onChange((v) => {
            this.kind = v as FileKind;
            if (!this.filenameTouched) {
              this.filename = `Untitled.${this.kind}`;
              if (this.filenameInput) this.filenameInput.value = this.filename;
            }
          });
      });

    new Setting(contentEl).setName("Filename").addText((t) => {
      t.setValue(this.filename).onChange((v) => {
        this.filename = v;
        this.filenameTouched = true;
      });
      this.filenameInput = t.inputEl;
      t.inputEl.style.width = "100%";
    });

    contentEl.createEl("label", {
      text: "Describe what Claude should create:",
      cls: "docx-claude-create-label",
    });
    const ta = contentEl.createEl("textarea", {
      cls: "docx-claude-create-textarea",
      attr: {
        rows: "8",
        placeholder:
          "e.g. A two-page brief on the Q3 product roadmap, with a heading, three bulleted priorities, and a small risks table.",
      },
    });
    ta.addEventListener("input", () => {
      this.description = ta.value;
    });

    this.statusEl = contentEl.createDiv({ cls: "docx-claude-create-status" });

    const actions = contentEl.createDiv({ cls: "docx-claude-create-actions" });
    const cancel = actions.createEl("button", { text: "Cancel" });
    cancel.addEventListener("click", () => this.close());

    this.submitBtn = actions.createEl("button", {
      text: "Create",
      cls: "mod-cta",
    }) as HTMLButtonElement;
    this.submitBtn.addEventListener("click", () => void this.submit());
  }

  onClose(): void {
    this.contentEl.empty();
  }

  private setStatus(text: string): void {
    if (this.statusEl) this.statusEl.setText(text);
  }

  private setBusy(busy: boolean): void {
    this.busy = busy;
    if (this.submitBtn) this.submitBtn.disabled = busy;
  }

  private resolveTargetFolder(): string {
    const active = this.app.workspace.getActiveFile();
    if (active && active.parent instanceof TFolder) return active.parent.path;
    return "";
  }

  private async pickAvailablePath(folder: string, baseName: string): Promise<string> {
    const joined = folder ? `${folder}/${baseName}` : baseName;
    const normalized = normalizePath(joined);
    if (!(await this.app.vault.adapter.exists(normalized))) return normalized;

    const dot = baseName.lastIndexOf(".");
    const stem = dot === -1 ? baseName : baseName.slice(0, dot);
    const ext = dot === -1 ? "" : baseName.slice(dot);
    for (let i = 1; i < 1000; i++) {
      const candidate = `${stem} (${i})${ext}`;
      const path = normalizePath(folder ? `${folder}/${candidate}` : candidate);
      if (!(await this.app.vault.adapter.exists(path))) return path;
    }
    throw new Error("Could not find an available filename.");
  }

  private normalizeFilename(raw: string, kind: FileKind): string {
    let name = raw.trim().replace(ILLEGAL_FILENAME_CHARS, " ");
    if (!name) name = "Untitled";
    const expectedExt = `.${kind}`;
    if (!name.toLowerCase().endsWith(expectedExt)) name += expectedExt;
    return name;
  }

  private async submit(): Promise<void> {
    if (this.busy) return;
    if (!this.description.trim()) {
      new Notice("Describe what to create first.");
      return;
    }
    this.setBusy(true);
    this.setStatus("Asking Claude…");
    try {
      const system = systemPromptFor(this.kind);
      const raw = await this.bridge.run(system, this.description.trim(), {
        timeoutMs: 180_000,
      });

      this.setStatus("Parsing response…");
      let parsed: unknown;
      try {
        parsed = extractJsonFromResponse(raw);
      } catch (e) {
        throw new Error(
          `Claude returned non-JSON output. ${
            e instanceof Error ? e.message : String(e)
          }`,
        );
      }
      const spec = validateSpec(this.kind, parsed);

      this.setStatus("Building file…");
      const buf = await buildFile(spec);

      const folder = this.resolveTargetFolder();
      const filename = this.normalizeFilename(this.filename, this.kind);
      const path = await this.pickAvailablePath(folder, filename);

      const file = await this.app.vault.createBinary(path, buf);
      new Notice(`Created ${file.path}`);
      this.close();
      await this.openFile(file);
    } catch (e) {
      const msg =
        e instanceof ClaudeBridgeError
          ? e.message
          : e instanceof SpecValidationError
            ? `Claude's JSON didn't match the expected shape: ${e.message}`
            : e instanceof Error
              ? e.message
              : String(e);
      this.setStatus(`Error: ${msg}`);
      new Notice(msg, 8000);
    } finally {
      this.setBusy(false);
    }
  }
}
