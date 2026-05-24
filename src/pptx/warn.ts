export function warn(scope: string, err: unknown, ctx?: Record<string, unknown>): void {
  console.warn(`[pptx:${scope}]`, err, ctx ?? {});
}
