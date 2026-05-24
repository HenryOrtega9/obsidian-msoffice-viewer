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
export async function renderPicInto(
  pic: Element,
  box: Box,
  rels: Map<string, string>,
  zip: JSZip,
  parent: HTMLElement,
  objectUrls: string[],
): Promise<void> {
  const el = parent.createDiv({ cls: "docx-claude-pptx-image" });
  setBox(el, box);

  const blipFill = directChild(pic, NS.p, "blipFill");
  const blip = blipFill ? directChild(blipFill, NS.a, "blip") : null;
  const embed = relId(blip);
  const path = embed ? rels.get(embed) : undefined;
  if (!path) {
    renderPlaceholder(el, "");
    return;
  }
  const ext = (path.split(".").pop() ?? "").toLowerCase();
  const mime = MIME_BY_EXT[ext];
  if (!mime) {
    renderPlaceholder(el, ext);
    return;
  }
  try {
    const file = zip.file(path);
    if (!file) {
      renderPlaceholder(el, ext);
      return;
    }
    const data = await file.async("uint8array");
    const url = URL.createObjectURL(new Blob([data as BlobPart], { type: mime }));
    objectUrls.push(url);
    const img = el.createEl("img", { cls: "docx-claude-pptx-img" });
    img.src = url;
  } catch (e) {
    warn("image-render", e, { path });
    renderPlaceholder(el, ext);
  }
}

function renderPlaceholder(el: HTMLElement, ext: string): void {
  el.addClass("docx-claude-pptx-image-placeholder");
  el.setText(ext ? `[${ext.toUpperCase()} image]` : "[image]");
}
