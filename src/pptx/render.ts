import type { Chart } from "chart.js";
import { loadPresentation } from "./presentation";
import { loadPresentationTheme, resolveEffectiveClrMap, type PptxTheme } from "./themes";
import { slideScaleFor } from "./geometry";
import { renderSlide } from "./slide";
import { warn } from "./warn";

const TARGET_SLIDE_WIDTH = 1280;

export interface NativeRenderOpts {
  isStale?: () => boolean;
}

export interface NativeRenderResult {
  slideCount: number;
  objectUrls: string[]; // populated from Phase 2 (images)
  charts: Chart[]; // populated from Phase 4 (charts)
}

// Render a deck natively into `parent`. Throws when the file is not a slide
// package or no slide could be rendered, so the caller falls through to the
// next tier (pptxviewjs, then the text card).
export async function renderDeckNatively(
  buf: ArrayBuffer,
  parent: HTMLElement,
  opts: NativeRenderOpts = {},
): Promise<NativeRenderResult> {
  const pkg = await loadPresentation(buf);
  if (!pkg || pkg.slides.length === 0) {
    throw new Error("Not a renderable slide package");
  }

  const scale = slideScaleFor(pkg.slideSize.cx, pkg.slideSize.cy, TARGET_SLIDE_WIDTH);
  const stage = parent.createDiv({ cls: "docx-claude-pdf-stage" });

  const objectUrls: string[] = [];
  const charts: Chart[] = [];
  const themeCache = new Map<string, PptxTheme | null>();
  let rendered = 0;

  for (const slideRef of pkg.slides) {
    if (opts.isStale?.()) break;
    const theme = await themeFor(pkg.zip, slideRef.themePath, themeCache);
    if (opts.isStale?.()) break;
    const clrMap = resolveEffectiveClrMap(slideRef.masterDoc, slideRef.layoutDoc, slideRef.slideDoc);
    try {
      await renderSlide(slideRef, theme, clrMap, scale, stage, pkg.zip, objectUrls, charts);
      rendered++;
    } catch (e) {
      warn("slide-render", e, { slide: slideRef.slidePath });
    }
  }

  if (rendered === 0) {
    throw new Error("Native renderer produced no slides");
  }
  return { slideCount: rendered, objectUrls, charts };
}

async function themeFor(
  zip: Parameters<typeof loadPresentationTheme>[0],
  path: string | null,
  cache: Map<string, PptxTheme | null>,
): Promise<PptxTheme | null> {
  const key = path ?? "";
  if (cache.has(key)) return cache.get(key) ?? null;
  const theme = await loadPresentationTheme(zip, path);
  cache.set(key, theme);
  return theme;
}
