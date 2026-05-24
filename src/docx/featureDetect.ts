import JSZip from "jszip";
import { warn } from "./warn";

export interface DocxComplexity {
  charts: boolean;
  smartArt: boolean;
  ole: boolean;
  omml: boolean;
  // True when the document contains parts docx-preview cannot render and that
  // need the LibreOffice PDF route. OMML is intentionally excluded (soft signal).
  forcePdf: boolean;
}

const CHART_RE = /^word\/charts\/chart\d+\.xml$/;
const SMARTART_RE = /^word\/(glossary\/)?diagrams\/data\d+\.xml$/;
const OLE_RE = /^word\/embeddings\//;

// Cheap two-tier scan over a single unzip: filename listing for charts /
// SmartArt / OLE, plus one document.xml text read for OMML.
export async function detectComplexFeatures(buf: ArrayBuffer): Promise<DocxComplexity> {
  const empty: DocxComplexity = {
    charts: false,
    smartArt: false,
    ole: false,
    omml: false,
    forcePdf: false,
  };
  try {
    const zip = await JSZip.loadAsync(buf);
    const names = Object.keys(zip.files);

    const charts = names.some((n) => CHART_RE.test(n));
    const smartArt = names.some((n) => SMARTART_RE.test(n));
    // Any embedded part that isn't a chart-owned workbook counts as OLE.
    const ole = names.some(
      (n) => OLE_RE.test(n) && !/Microsoft_Excel|chart/i.test(n),
    );

    let omml = false;
    const docXml = zip.file("word/document.xml");
    if (docXml) {
      const text = await docXml.async("text");
      omml = text.includes("<m:oMath");
    }

    return {
      charts,
      smartArt,
      ole,
      omml,
      forcePdf: charts || smartArt || ole,
    };
  } catch (e) {
    warn("feature-detect", e);
    return empty;
  }
}
