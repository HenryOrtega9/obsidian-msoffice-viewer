import { App, Modal } from "obsidian";
import type { StructuralLocator } from "./types";

export class EditInstructionModal extends Modal {
  private instruction = "";
  private resolved = false;

  constructor(
    app: App,
    private locator: StructuralLocator,
    private onSubmit: (instruction: string) => void,
    private onCancel: () => void,
  ) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h3", { text: "Ask Claude to edit selection" });

    const preview = contentEl.createDiv({ cls: "docx-claude-modal-selection" });
    preview.setText(this.locator.selectedText || "(empty selection)");

    const textarea = contentEl.createEl("textarea", {
      cls: "docx-claude-modal-instruction",
      attr: { placeholder: "Describe the rewrite you want..." },
    });
    textarea.addEventListener("input", () => {
      this.instruction = textarea.value;
    });
    textarea.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" && (ev.metaKey || ev.ctrlKey)) {
        ev.preventDefault();
        this.submit();
      }
    });
    textarea.focus();

    const buttons = contentEl.createDiv({ cls: "docx-claude-modal-buttons" });
    const cancel = buttons.createEl("button", { text: "Cancel" });
    cancel.addEventListener("click", () => this.close());
    const submit = buttons.createEl("button", {
      text: "Send to Claude",
      cls: "mod-cta",
    });
    submit.addEventListener("click", () => this.submit());
  }

  private submit(): void {
    const trimmed = this.instruction.trim();
    if (!trimmed) return;
    this.resolved = true;
    this.onSubmit(trimmed);
    this.close();
  }

  onClose(): void {
    this.contentEl.empty();
    if (!this.resolved) this.onCancel();
  }
}
