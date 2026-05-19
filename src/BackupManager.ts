import { App, TFile, normalizePath } from "obsidian";

export class BackupManager {
  private sessionBackups = new Set<string>();
  private roundtripCounts = new Map<string, number>();

  constructor(
    private app: App,
    private pluginDir: string,
    private maxHistory: number,
  ) {}

  setMaxHistory(n: number): void {
    this.maxHistory = n;
  }

  async snapshot(file: TFile): Promise<void> {
    await this.ensureSessionBackup(file);
    await this.archive(file);
  }

  private async ensureSessionBackup(file: TFile): Promise<void> {
    if (this.sessionBackups.has(file.path)) return;
    const bakPath = `${file.path}.bak`;
    const exists = await this.app.vault.adapter.exists(bakPath);
    if (!exists) {
      const buf = await this.app.vault.readBinary(file);
      await this.app.vault.adapter.writeBinary(bakPath, buf);
    }
    this.sessionBackups.add(file.path);
  }

  private async archive(file: TFile): Promise<void> {
    if (this.maxHistory <= 0) return;
    const dir = normalizePath(`${this.pluginDir}/history/${file.basename}`);
    await this.ensureDir(dir);
    const ts = new Date().toISOString().replace(/[:.]/g, "-");
    const target = `${dir}/${ts}.docx`;
    const buf = await this.app.vault.readBinary(file);
    await this.app.vault.adapter.writeBinary(target, buf);
    await this.prune(dir);
  }

  private async ensureDir(dir: string): Promise<void> {
    const exists = await this.app.vault.adapter.exists(dir);
    if (!exists) {
      await this.app.vault.adapter.mkdir(dir);
    }
  }

  private async prune(dir: string): Promise<void> {
    const listing = await this.app.vault.adapter.list(dir);
    const files = listing.files
      .filter((f) => f.endsWith(".docx"))
      .sort();
    while (files.length > this.maxHistory) {
      const oldest = files.shift();
      if (oldest) await this.app.vault.adapter.remove(oldest);
    }
  }

  recordRoundtrip(filePath: string): number {
    const next = (this.roundtripCounts.get(filePath) ?? 0) + 1;
    this.roundtripCounts.set(filePath, next);
    return next;
  }

  getRoundtripCount(filePath: string): number {
    return this.roundtripCounts.get(filePath) ?? 0;
  }
}
