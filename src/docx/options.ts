import type { Options } from "docx-preview";

// docx-preview render options. Kept here so the view stays thin and the option
// set is easy to tune. docx-preview resolves theme colors/fonts and embedded
// fonts itself, so there's no custom parsing to do for those.
export function buildDocxOptions(): Partial<Options> {
  return {
    className: "docx-claude",
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: true,
    experimental: true,
    // Object URLs can be flaky for fonts/images in Electron; base64 is reliable
    // for an offline desktop viewer.
    useBase64URL: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    renderComments: true,
    renderChanges: true,
  };
}
