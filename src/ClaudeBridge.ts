import { App } from "obsidian";

const CLAUDE_PLUGIN_ID = "claude-cli-chat";

interface RunHeadlessOpts {
  timeoutMs?: number;
  cwd?: string;
}

interface ClaudePluginLike {
  runHeadlessPrompt?: (
    systemPrompt: string,
    userPrompt: string,
    opts?: RunHeadlessOpts,
  ) => Promise<string>;
}

interface PluginRegistryLike {
  plugins?: { plugins?: Record<string, unknown>; enabledPlugins?: Set<string> };
}

export class ClaudeBridgeError extends Error {}

export class ClaudeBridge {
  constructor(private readonly app: App) {}

  private resolvePlugin(): ClaudePluginLike {
    const registry = (this.app as unknown as PluginRegistryLike).plugins;
    const plugin = registry?.plugins?.[CLAUDE_PLUGIN_ID] as ClaudePluginLike | undefined;
    if (!plugin) {
      throw new ClaudeBridgeError(
        "Claude plugin (claude-cli-chat) is not installed or not enabled. Install/enable it to use this feature.",
      );
    }
    if (typeof plugin.runHeadlessPrompt !== "function") {
      throw new ClaudeBridgeError(
        "Installed claude-cli-chat is too old: it does not expose runHeadlessPrompt. Update it.",
      );
    }
    return plugin;
  }

  async run(systemPrompt: string, userPrompt: string, opts?: RunHeadlessOpts): Promise<string> {
    const plugin = this.resolvePlugin();
    const timeoutMs = opts?.timeoutMs ?? 180_000;
    // Local timer so we don't hang forever if the embedded plugin ignores
    // timeoutMs or never resolves (e.g., subprocess deadlock).
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(
        () => reject(new ClaudeBridgeError(`Claude call timed out after ${timeoutMs}ms.`)),
        timeoutMs,
      );
    });
    try {
      const call = plugin.runHeadlessPrompt!(systemPrompt, userPrompt, opts);
      // If the timeout wins the race, the losing promise keeps running with no
      // handler attached; absorb its eventual rejection so it doesn't surface
      // as an unhandled rejection.
      call.catch(() => {});
      return await Promise.race([call, timeout]);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }
}
