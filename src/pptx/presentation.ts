import JSZip from "jszip";
import {
  NS,
  readXml,
  readRelsDetailed,
  relTargetByType,
  relsPathFor,
  relId,
  firstChildNS,
  directChild,
  childrenNS,
  intAttr,
} from "./ooxml";
import { warn } from "./warn";

export interface SlideRef {
  slidePath: string;
  slideDoc: Document;
  layoutDoc: Document | null;
  masterDoc: Document | null;
  themePath: string | null;
  rels: Map<string, string>; // id -> resolved zip path (images, charts)
  layoutRels: Map<string, string>; // rels of the layout part (for its images)
  masterRels: Map<string, string>; // rels of the master part
}

export interface PresentationPackage {
  zip: JSZip;
  slideSize: { cx: number; cy: number };
  slides: SlideRef[];
}

// Default 16:9 slide size used only if presentation.xml omits sldSz. Modern
// decks are widescreen, so fall back to 16:9 (12192000 x 6858000) rather than
// the legacy 4:3 (9144000 x 6858000).
const DEFAULT_SLIDE_CX = 12192000;
const DEFAULT_SLIDE_CY = 6858000;

// Open the package and resolve slides in authoritative presentation order
// (p:sldIdLst -> rels), each carrying its layout/master/theme chain. Returns
// null if this is not a slide-bearing pptx.
export async function loadPresentation(buf: ArrayBuffer): Promise<PresentationPackage | null> {
  try {
    const zip = await JSZip.loadAsync(buf);
    if (!Object.keys(zip.files).some((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))) {
      return null;
    }
    const presDoc = await readXml(zip, "ppt/presentation.xml");
    if (!presDoc) return null;

    const sldSz = firstChildNS(presDoc.documentElement, NS.p, "sldSz");
    const slideSize = {
      cx: intAttr(sldSz, "cx", DEFAULT_SLIDE_CX),
      cy: intAttr(sldSz, "cy", DEFAULT_SLIDE_CY),
    };

    const presRels = await readRelsDetailed(zip, relsPathFor("ppt/presentation.xml"), "ppt");
    const presRelMap = new Map(presRels.map((r) => [r.id, r.target]));

    const slides: SlideRef[] = [];
    const sldIdLst = firstChildNS(presDoc.documentElement, NS.p, "sldIdLst");
    if (sldIdLst) {
      for (const sldId of childrenNS(sldIdLst, NS.p, "sldId")) {
        const rid = relId(sldId);
        const slidePath = rid ? presRelMap.get(rid) : undefined;
        if (!slidePath) continue;
        const ref = await loadSlideRef(zip, slidePath);
        if (ref) slides.push(ref);
      }
    }

    return { zip, slideSize, slides };
  } catch (e) {
    warn("presentation-load", e);
    return null;
  }
}

async function loadSlideRef(zip: JSZip, slidePath: string): Promise<SlideRef | null> {
  const slideDoc = await readXml(zip, slidePath);
  if (!slideDoc) return null;

  const slideRels = await readRelsDetailed(zip, relsPathFor(slidePath), dirOfSlide(slidePath));
  const rels = new Map(slideRels.map((r) => [r.id, r.target]));

  let layoutDoc: Document | null = null;
  let masterDoc: Document | null = null;
  let themePath: string | null = null;
  let layoutRelMap = new Map<string, string>();
  let masterRelMap = new Map<string, string>();

  const layoutPath = relTargetByType(slideRels, "/slideLayout");
  if (layoutPath) {
    layoutDoc = await readXml(zip, layoutPath);
    const layoutRels = await readRelsDetailed(zip, relsPathFor(layoutPath), dirOfSlide(layoutPath));
    layoutRelMap = new Map(layoutRels.map((r) => [r.id, r.target]));
    const masterPath = relTargetByType(layoutRels, "/slideMaster");
    if (masterPath) {
      masterDoc = await readXml(zip, masterPath);
      const masterRels = await readRelsDetailed(zip, relsPathFor(masterPath), dirOfSlide(masterPath));
      masterRelMap = new Map(masterRels.map((r) => [r.id, r.target]));
      themePath = relTargetByType(masterRels, "/theme");
    }
  }

  return {
    slidePath,
    slideDoc,
    layoutDoc,
    masterDoc,
    themePath,
    rels,
    layoutRels: layoutRelMap,
    masterRels: masterRelMap,
  };
}

function dirOfSlide(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

// Re-export so consumers can find the spTree consistently.
export function findSpTree(doc: Document): Element | null {
  const cSld = doc.getElementsByTagNameNS(NS.p, "cSld")[0];
  if (!cSld) return null;
  return directChild(cSld, NS.p, "spTree");
}
