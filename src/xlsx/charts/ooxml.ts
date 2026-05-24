import type JSZip from "jszip";

export const NS = {
  rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  pkgRel: "http://schemas.openxmlformats.org/package/2006/relationships",
  ssml: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
};

export async function readXml(zip: JSZip, path: string): Promise<Document | null> {
  const file = zip.file(path);
  if (!file) return null;
  const text = await file.async("text");
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) return null;
  return doc;
}

// Parse a .rels file into a map of relationship id → resolved zip path.
// basePath is the directory of the part owning the rels (e.g. "xl/worksheets").
export async function readRels(
  zip: JSZip,
  relsPath: string,
  basePath: string,
): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const doc = await readXml(zip, relsPath);
  if (!doc) return out;
  const rels = doc.getElementsByTagNameNS(NS.pkgRel, "Relationship");
  for (const rel of Array.from(rels)) {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (!id || !target) continue;
    out.set(id, resolvePath(basePath, target));
  }
  return out;
}

// Resolve a relationship Target (often "../drawings/drawing1.xml") against the
// owning part's directory into a normalized zip path.
export function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const baseParts = baseDir.split("/").filter(Boolean);
  const targetParts = target.split("/");
  for (const part of targetParts) {
    if (part === "..") baseParts.pop();
    else if (part === "." || part === "") continue;
    else baseParts.push(part);
  }
  return baseParts.join("/");
}

export function dirOf(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

// r:id attribute lookup that tolerates namespace-prefixed attributes across
// DOMParser implementations.
export function relId(el: Element): string | null {
  return el.getAttributeNS(NS.rel, "id") ?? el.getAttribute("r:id");
}

export function firstChildNS(parent: Element, ns: string, local: string): Element | null {
  const found = parent.getElementsByTagNameNS(ns, local);
  return found.length > 0 ? found[0] : null;
}

export function textOfChildNS(parent: Element, ns: string, local: string): string {
  const el = firstChildNS(parent, ns, local);
  return el?.textContent ?? "";
}
