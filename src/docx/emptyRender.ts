// Detect when docx-preview produced no useful output (it can resolve a file
// without throwing yet render an empty wrapper). Treats image/table/canvas-only
// pages as non-empty so we don't misfire the PDF fallback on figure-heavy docs.
export function isRenderEmpty(renderEl: HTMLElement): boolean {
  if (renderEl.childElementCount === 0) return true;
  const wrapper = renderEl.querySelector<HTMLElement>(".docx-claude-wrapper, .docx-wrapper");
  if (!wrapper) return true;
  const hasText = (wrapper.textContent ?? "").trim().length > 0;
  if (hasText) return false;
  return wrapper.querySelector("img, svg, table, canvas") === null;
}
