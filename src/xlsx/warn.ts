export function warn(scope: string, err: unknown, ctx?: Record<string, unknown>): void {
  console.warn(`[xlsx:${scope}]`, err, ctx ?? {});
}
