import JSZip from "jszip";
import { DEFAULT_THEME_RGB } from "./colors";
import { warn } from "./warn";

const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

// OOXML clrScheme children appear in file order dk1, lt1, dk2, lt2, accent1..6,
// hlink, folHlink. Excel's <color theme="n"> indexes a DIFFERENT order that
// swaps the first two pairs: 0=lt1(Background1), 1=dk1(Text1), 2=lt2, 3=dk2,
// then accent1..6 (4..9), hlink (10), folHlink (11). We key on each child's
// localName (not file position), so this maps each scheme slot to the Excel
// theme index ExcelJS reports on cells. (Getting 0/1 backwards renders default
// black Text1 as white — the dk1/lt1 slots must not be swapped.)
const FILE_TO_EXCEL_INDEX: Record<string, number> = {
  lt1: 0,
  dk1: 1,
  lt2: 2,
  dk2: 3,
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  hlink: 10,
  folHlink: 11,
};

export async function loadWorkbookTheme(buf: ArrayBuffer): Promise<string[] | null> {
  try {
    const zip = await JSZip.loadAsync(buf);
    // theme1.xml is overwhelmingly the active theme; if it's missing, fall back
    // to the first themeN.xml present.
    let themeFile = zip.file("xl/theme/theme1.xml");
    if (!themeFile) {
      const candidates = Object.keys(zip.files).filter((p) =>
        /^xl\/theme\/theme\d+\.xml$/.test(p),
      );
      if (candidates.length === 0) return null;
      themeFile = zip.file(candidates[0]) ?? null;
    }
    if (!themeFile) return null;
    const xml = await themeFile.async("text");
    return parseThemeXml(xml);
  } catch (e) {
    warn("theme-load", e);
    return null;
  }
}

function parseThemeXml(xml: string): string[] | null {
  try {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const parserError = doc.getElementsByTagName("parsererror")[0];
    if (parserError) {
      warn("theme-parse", parserError.textContent);
      return null;
    }
    const schemes = doc.getElementsByTagNameNS(DRAWINGML_NS, "clrScheme");
    if (schemes.length === 0) return null;
    const scheme = schemes[0];
    // Start from defaults; only override slots we resolve.
    const out = [...DEFAULT_THEME_RGB];

    for (const node of Array.from(scheme.childNodes)) {
      if (node.nodeType !== 1) continue; // element
      const el = node as Element;
      const localName = el.localName;
      const excelIdx = FILE_TO_EXCEL_INDEX[localName];
      if (excelIdx === undefined) continue;
      const hex = extractColorFromSchemeEntry(el);
      if (hex) out[excelIdx] = hex;
    }
    return out;
  } catch (e) {
    warn("theme-parse", e);
    return null;
  }
}

// Each clrScheme child contains exactly one color child: <a:srgbClr val="HEX"/>
// or <a:sysClr lastClr="HEX" val="windowText"/>. Other forms (schemeClr refs
// nested here) are rare for clrScheme itself — warn and skip.
function extractColorFromSchemeEntry(el: Element): string | null {
  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType !== 1) continue;
    const colorEl = child as Element;
    if (colorEl.namespaceURI !== DRAWINGML_NS) continue;
    if (colorEl.localName === "srgbClr") {
      const val = colorEl.getAttribute("val");
      if (val && /^[0-9A-Fa-f]{6}$/.test(val)) return val.toUpperCase();
    }
    if (colorEl.localName === "sysClr") {
      const last = colorEl.getAttribute("lastClr");
      if (last && /^[0-9A-Fa-f]{6}$/.test(last)) return last.toUpperCase();
    }
    if (colorEl.localName === "schemeClr") {
      warn("theme-schemeClr-ref", null, { entry: el.localName });
    }
  }
  return null;
}
