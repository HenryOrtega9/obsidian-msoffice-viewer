import { App, PluginSettingTab, Setting } from "obsidian";
import type MsOfficeViewerPlugin from "./main";

export const MIN_ZOOM = 0.25;
export const MAX_ZOOM = 4.0;
export const ZOOM_STEP = 0.1;
export const DEFAULT_ZOOM = 1.0;

export interface DocxPreviewSettings {
  defaultZoom: number;
}

export const DEFAULT_SETTINGS: DocxPreviewSettings = {
  defaultZoom: DEFAULT_ZOOM,
};

export function clampZoom(z: number): number {
  if (!Number.isFinite(z)) return DEFAULT_ZOOM;
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z));
}

export class DocxPreviewSettingTab extends PluginSettingTab {
  plugin: MsOfficeViewerPlugin;

  constructor(app: App, plugin: MsOfficeViewerPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    new Setting(containerEl)
      .setName("Default zoom")
      .setDesc(`Starting zoom level for newly opened Office files. ${Math.round(MIN_ZOOM * 100)}%–${Math.round(MAX_ZOOM * 100)}%.`)
      .addText((text) => {
        text
          .setPlaceholder("100")
          .setValue(String(Math.round(this.plugin.settings.defaultZoom * 100)));
        // Clamp, persist, and reflect ONLY on blur. Doing it in onChange fires
        // per keystroke: typing "50" clamps the intermediate "5" to 25 and
        // rewrites the field mid-typing, hijacking the input.
        const commit = async (): Promise<void> => {
          const pct = Number.parseFloat(text.inputEl.value);
          if (!Number.isFinite(pct)) {
            text.inputEl.value = String(Math.round(this.plugin.settings.defaultZoom * 100));
            return;
          }
          const next = clampZoom(pct / 100);
          this.plugin.settings.defaultZoom = next;
          await this.plugin.saveSettings();
          text.inputEl.value = String(Math.round(next * 100));
        };
        text.inputEl.addEventListener("blur", () => void commit());
        text.inputEl.addEventListener("keydown", (ev) => {
          if (ev.key === "Enter") text.inputEl.blur();
        });
      });
  }
}
