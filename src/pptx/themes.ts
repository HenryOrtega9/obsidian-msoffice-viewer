import type JSZip from "jszip";
import { NS, readXml, elementChildren, directChild, firstChildNS } from "./ooxml";
import { warn } from "./warn";

// Theme colors keyed by scheme name. PresentationML references theme colors by
// NAME (dk1, lt1, accent1, ...) via a:schemeClr, so no Excel-style index swap
// is needed here (unlike src/xlsx/themes.ts).
export interface PptxTheme {
  scheme: Record<string, string>; // name -> "RRGGBB"
  majorFont?: string; // a:fontScheme/a:majorFont latin typeface (+mj-lt)
  minorFont?: string; // a:fontScheme/a:minorFont latin typeface (+mn-lt)
}

// p:clrMap maps logical roles (bg1, tx1, bg2, tx2, ...) onto scheme names.
export type ClrMap = Record<string, string>;

// Default Office theme palette, by scheme name.
export const DEFAULT_SCHEME: Record<string, string> = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  hlink: "0563C1",
  folHlink: "954F72",
};

const DEFAULT_CLR_MAP: ClrMap = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

export async function loadPresentationTheme(
  zip: JSZip,
  themePath: string | null,
): Promise<PptxTheme | null> {
  try {
    let path = themePath;
    if (!path) {
      const candidate = Object.keys(zip.files).find((p) =>
        /^ppt\/theme\/theme\d+\.xml$/.test(p),
      );
      if (!candidate) return null;
      path = candidate;
    }
    const doc = await readXml(zip, path);
    if (!doc) return null;
    return parseThemeDoc(doc);
  } catch (e) {
    warn("theme-load", e, { themePath });
    return null;
  }
}

function parseThemeDoc(doc: Document): PptxTheme | null {
  const schemes = doc.getElementsByTagNameNS(NS.a, "clrScheme");
  if (schemes.length === 0) return null;
  const scheme = schemes[0];
  const out: Record<string, string> = { ...DEFAULT_SCHEME };
  for (const node of elementChildren(scheme)) {
    const hex = extractSchemeColor(node);
    if (hex) out[node.localName] = hex;
  }
  return { scheme: out, majorFont: fontOf(doc, "majorFont"), minorFont: fontOf(doc, "minorFont") };
}

function fontOf(doc: Document, which: string): string | undefined {
  const fs = doc.getElementsByTagNameNS(NS.a, "fontScheme")[0];
  if (!fs) return undefined;
  const font = directChild(fs, NS.a, which);
  const latin = font ? directChild(font, NS.a, "latin") : null;
  return latin?.getAttribute("typeface") || undefined;
}

// Each clrScheme child (dk1, lt1, ...) wraps one a:srgbClr or a:sysClr.
function extractSchemeColor(entry: Element): string | null {
  const srgb = directChild(entry, NS.a, "srgbClr");
  if (srgb) {
    const v = srgb.getAttribute("val");
    if (v && /^[0-9A-Fa-f]{6}$/.test(v)) return v.toUpperCase();
  }
  const sys = directChild(entry, NS.a, "sysClr");
  if (sys) {
    const last = sys.getAttribute("lastClr");
    if (last && /^[0-9A-Fa-f]{6}$/.test(last)) return last.toUpperCase();
  }
  return null;
}

export function parseClrMap(masterDoc: Document | null): ClrMap {
  const map: ClrMap = { ...DEFAULT_CLR_MAP };
  if (!masterDoc) return map;
  const cm = firstChildNS(masterDoc.documentElement, NS.p, "clrMap");
  if (cm) readClrMapAttrs(cm, map);
  return map;
}

function readClrMapAttrs(el: Element, into: ClrMap): void {
  for (const role of Object.keys(into)) {
    const v = el.getAttribute(role);
    if (v) into[role] = v;
  }
}

// The effective color map is the master's clrMap, unless the slide (then the
// layout) carries a p:clrMapOvr with an a:overrideClrMapping.
export function resolveEffectiveClrMap(
  masterDoc: Document | null,
  layoutDoc: Document | null,
  slideDoc: Document | null,
): ClrMap {
  const base = parseClrMap(masterDoc);
  for (const doc of [slideDoc, layoutDoc]) {
    if (!doc) continue;
    const ovr = firstChildNS(doc.documentElement, NS.p, "clrMapOvr");
    if (!ovr) continue;
    const override = firstChildNS(ovr, NS.a, "overrideClrMapping");
    if (override) {
      const map: ClrMap = { ...DEFAULT_CLR_MAP };
      readClrMapAttrs(override, map);
      return map;
    }
    // a:masterClrMapping (or empty) means inherit the master's map.
    return base;
  }
  return base;
}
