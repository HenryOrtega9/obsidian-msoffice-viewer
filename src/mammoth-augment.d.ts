// Augment mammoth's typings: convertToMarkdown exists at runtime but is
// missing from the bundled .d.ts.
declare module "mammoth" {
  type AnyInput =
    | { path: string }
    | { buffer: Buffer }
    | { arrayBuffer: ArrayBuffer };

  interface MarkdownResult {
    value: string;
    messages: Array<{ type: string; message: string }>;
  }

  export function convertToMarkdown(input: AnyInput): Promise<MarkdownResult>;
}
