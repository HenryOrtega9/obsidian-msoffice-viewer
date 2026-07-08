import type JSZip from "jszip";
import { NS, directChild, relId } from "./ooxml";
import type { Box } from "./geometry";
import { warn } from "./warn";

const MIME_BY_EXT: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  tif: "image/tiff",
  tiff: "image/tiff",
};

function setBox(el: HTMLElement, box: Box): void {
  el.style.left = `${box.left}px`;
  el.style.top = `${box.top}px`;
  el.style.width = `${box.width}px`;
  el.style.height = `${box.height}px`;
}

// Render a p:pic into an absolutely-positioned <img>, reading bytes from the
// package and minting an object URL (pushed to `objectUrls` for later revoke).
// Vector formats the browser can't show (EMF/WMF) get a placeholder box.
// Returns the holder element so the caller can apply rotate/flip transforms.
export async function renderPicInto(
  pic: Element,
  box: Box,
  rels: Map<string, string>,
  zip: JSZip,
  parent: HTMLElement,
  objectUrls: string[],
): Promise<HTMLElement> {
  const el = parent.createDiv({ cls: "docx-claude-pptx-image" });
  setBox(el, box);

  const blipFill = directChild(pic, NS.p, "blipFill");
  const blip = blipFill ? directChild(blipFill, NS.a, "blip") : null;
  const embed = relId(blip);
  const path = embed ? rels.get(embed) : undefined;
  if (!path) {
    renderPlaceholder(el, "");
    return el;
  }
  const ext = (path.split(".").pop() ?? "").toLowerCase();
  const mime = MIME_BY_EXT[ext];
  if (!mime) {
    renderPlaceholder(el, ext);
    return el;
  }
  try {
    const file = zip.file(path);
    if (!file) {
      renderPlaceholder(el, ext);
      return el;
    }
    const data = await file.async("uint8array");
    const url = URL.createObjectURL(new Blob([data as BlobPart], { type: mime }));
    objectUrls.push(url);
    const img = el.createEl("img", { cls: "docx-claude-pptx-img" });
    img.src = url;
    applySrcRect(el, img, blipFill);
  } catch (e) {
    warn("image-render", e, { path });
    renderPlaceholder(el, ext);
  }
  return el;
}

// a:srcRect crops the source picture before it fills the frame; l/t/r/b are
// crop-away percentages in 1000ths of a percent. Render by oversizing the img
// inside the clipping holder: visible fraction (1-l-r) horizontally maps to
// 100% of the holder, so the img is 1/(1-l-r) wide and shifted left by l.
function applySrcRect(holder: HTMLElement, img: HTMLImageElement, blipFill: Element | null): void {
  const srcRect = blipFill ? directChild(blipFill, NS.a, "srcRect") : null;
  if (!srcRect) return;
  const pct = (name: string): number => {
    const v = parseInt(srcRect.getAttribute(name) ?? "0", 10);
    return Number.isFinite(v) ? v / 100000 : 0;
  };
  const l = pct("l");
  const r = pct("r");
  const t = pct("t");
  const b = pct("b");
  const visW = 1 - l - r;
  const visH = 1 - t - b;
  if (visW <= 0 || visH <= 0 || (l === 0 && r === 0 && t === 0 && b === 0)) return;
  img.style.position = "absolute";
  img.style.width = `${(100 / visW).toFixed(3)}%`;
  img.style.height = `${(100 / visH).toFixed(3)}%`;
  img.style.left = `${(-(l / visW) * 100).toFixed(3)}%`;
  img.style.top = `${(-(t / visH) * 100).toFixed(3)}%`;
}

function renderPlaceholder(el: HTMLElement, ext: string): void {
  el.addClass("docx-claude-pptx-image-placeholder");
  el.setText(ext ? `[${ext.toUpperCase()} image]` : "[image]");
}
