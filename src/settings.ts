import { App, PluginSettingTab, Setting } from "obsidian";
import type DocxClaudePlugin from "./main";

export interface DocxClaudeSettings {
  pandocPath: string;
  maxHistory: number;
  warnAfterNRoundtrips: number;
}

export const DEFAULT_SETTINGS: DocxClaudeSettings = {
  pandocPath: "pandoc",
  maxHistory: 20,
  warnAfterNRoundtrips: 3,
};

export class DocxClaudeSettingTab extends PluginSettingTab {
  plugin: DocxClaudePlugin;

  constructor(app: App, plugin: DocxClaudePlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    new Setting(containerEl)
      .setName("Pandoc path")
      .setDesc(
        "Path to the pandoc binary used for markdown -> docx round-trip edits. Install with: brew install pandoc",
      )
      .addText((text) =>
        text
          .setPlaceholder("pandoc")
          .setValue(this.plugin.settings.pandocPath)
          .onChange(async (value) => {
            this.plugin.settings.pandocPath = value.trim() || "pandoc";
            await this.plugin.saveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Max history snapshots per file")
      .setDesc(
        "Older snapshots are pruned after this count. Stored under <plugin>/history/<basename>/.",
      )
      .addText((text) =>
        text
          .setValue(String(this.plugin.settings.maxHistory))
          .onChange(async (value) => {
            const n = Number.parseInt(value, 10);
            if (!Number.isNaN(n) && n >= 0) {
              this.plugin.settings.maxHistory = n;
              await this.plugin.saveSettings();
            }
          }),
      );

    new Setting(containerEl)
      .setName("Warn after N round-trip edits")
      .setDesc(
        "Show a drift warning in the preview after this many markdown round-trip edits in a session.",
      )
      .addText((text) =>
        text
          .setValue(String(this.plugin.settings.warnAfterNRoundtrips))
          .onChange(async (value) => {
            const n = Number.parseInt(value, 10);
            if (!Number.isNaN(n) && n >= 0) {
              this.plugin.settings.warnAfterNRoundtrips = n;
              await this.plugin.saveSettings();
            }
          }),
      );
  }
}
