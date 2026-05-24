import type ExcelJS from "exceljs";
import type { GridContext } from "./grid";
import { warn } from "./warn";
import { Box, EMU_PER_PX, anchorRangeToBox, type AnchorPoint } from "./geometry";

const MAX_IMAGES_PER_SHEET = 200;

interface AnchorLike {
  nativeCol?: number;
  nativeColOff?: number;
  nativeRow?: number;
  nativeRowOff?: number;
  col?: number;
  row?: number;
}

export interface ImageRenderResult {
  // Object URLs created via URL.createObjectURL — caller must revoke on unload.
  objectUrls: string[];
}

// Render embedded images into an absolute overlay layer above the grid table.
// The layer is created inside ctx.sheetWrapEl (position: relative), so images
// scroll in lockstep with the cells and share the grid's coordinate space.
export function renderSheetImages(
  ws: ExcelJS.Worksheet,
  wb: ExcelJS.Workbook,
  ctx: GridContext,
): ImageRenderResult {
  const objectUrls: string[] = [];
  let images: Array<{ imageId: string; range: unknown }>;
  try {
    images = ws.getImages() as Array<{ imageId: string; range: unknown }>;
  } catch (e) {
    warn("images-list", e);
    return { objectUrls };
  }
  if (!images || images.length === 0) return { objectUrls };

  const layer = ctx.sheetWrapEl.createDiv({ cls: "docx-claude-xlsx-image-layer" });

  let count = 0;
  for (const entry of images) {
    if (count >= MAX_IMAGES_PER_SHEET) {
      warn("images-cap", null, { cap: MAX_IMAGES_PER_SHEET, total: images.length });
      break;
    }
    try {
      const media = wb.getImage(parseInt(entry.imageId, 10));
      if (!media) continue;
      const box = anchorToBox(entry.range, ctx);
      if (!box) continue;

      const ext = media.extension;
      if (!isRenderableExtension(ext)) {
        appendPlaceholder(layer, box, ext);
        count++;
        continue;
      }

      const src = mediaToSrc(media, ext, objectUrls);
      if (!src) continue;

      const img = layer.createEl("img", { cls: "docx-claude-xlsx-image" });
      img.src = src;
      img.style.left = `${box.left}px`;
      img.style.top = `${box.top}px`;
      img.style.width = `${box.width}px`;
      img.style.height = `${box.height}px`;
      count++;
    } catch (e) {
      warn("image-render", e, { imageId: entry.imageId });
    }
  }

  return { objectUrls };
}

function isRenderableExtension(ext: string): boolean {
  return ext === "png" || ext === "jpeg" || ext === "gif";
}

function mediaToSrc(
  media: ExcelJS.Image,
  ext: string,
  objectUrls: string[],
): string | null {
  const mime = ext === "jpeg" ? "image/jpeg" : `image/${ext}`;
  if (media.base64) {
    // ExcelJS may hand back a bare base64 string or a full data URI.
    if (media.base64.startsWith("data:")) return media.base64;
    return `data:${mime};base64,${media.base64}`;
  }
  if (media.buffer) {
    const blob = new Blob([media.buffer], { type: mime });
    const url = URL.createObjectURL(blob);
    objectUrls.push(url);
    return url;
  }
  return null;
}

function toAnchorPoint(a: AnchorLike): AnchorPoint {
  return {
    col: a.nativeCol ?? Math.floor(a.col ?? 0),
    colOff: a.nativeColOff ?? 0,
    row: a.nativeRow ?? Math.floor(a.row ?? 0),
    rowOff: a.nativeRowOff ?? 0,
  };
}

function anchorToBox(range: unknown, ctx: GridContext): Box | null {
  if (!range || typeof range !== "object") return null;
  const r = range as { tl?: AnchorLike; br?: AnchorLike; ext?: { width?: number; height?: number } };
  if (!r.tl) return null;
  const from = toAnchorPoint(r.tl);
  const to = r.br ? toAnchorPoint(r.br) : null;
  const ext = r.ext ? { width: r.ext.width ?? 0, height: r.ext.height ?? 0 } : null;
  return anchorRangeToBox(from, to, ext, ctx);
}

function appendPlaceholder(layer: HTMLElement, box: Box, ext: string): void {
  const ph = layer.createDiv({ cls: "docx-claude-xlsx-image-placeholder" });
  ph.style.left = `${box.left}px`;
  ph.style.top = `${box.top}px`;
  ph.style.width = `${box.width}px`;
  ph.style.height = `${box.height}px`;
  ph.setText(`${ext.toUpperCase()} image not supported`);
}
