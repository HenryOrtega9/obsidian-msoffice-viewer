import type JSZip from "jszip";

// PresentationML namespaces. Mirrors src/xlsx/charts/ooxml.ts but adds the `p`
// (presentationml) namespace and drops the spreadsheet-only ones.
export const NS = {
  rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  pkgRel: "http://schemas.openxmlformats.org/package/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
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

export interface RelEntry {
  id: string;
  type: string;
  target: string; // resolved zip path (or external URL when TargetMode=External)
  external: boolean;
}

// Parse a .rels file into typed entries. PPTX needs Type-based lookups (a slide
// points at its layout, a layout at its master, a master at its theme) on top
// of the id->path map xlsx used.
export async function readRelsDetailed(
  zip: JSZip,
  relsPath: string,
  basePath: string,
): Promise<RelEntry[]> {
  const out: RelEntry[] = [];
  const doc = await readXml(zip, relsPath);
  if (!doc) return out;
  for (const rel of Array.from(doc.getElementsByTagNameNS(NS.pkgRel, "Relationship"))) {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (!id || !target) continue;
    const external = rel.getAttribute("TargetMode") === "External";
    out.push({
      id,
      type: rel.getAttribute("Type") ?? "",
      target: external ? target : resolvePath(basePath, target),
      external,
    });
  }
  return out;
}

// id -> resolved path map (used for r:embed image / r:id chart lookups).
export async function readRels(
  zip: JSZip,
  relsPath: string,
  basePath: string,
): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const rel of await readRelsDetailed(zip, relsPath, basePath)) {
    out.set(rel.id, rel.target);
  }
  return out;
}

export function relTargetByType(rels: RelEntry[], typeSuffix: string): string | null {
  const found = rels.find((r) => r.type.endsWith(typeSuffix));
  return found ? found.target : null;
}

// Resolve a relationship Target (often "../slideLayouts/slideLayout1.xml")
// against the owning part's directory into a normalized zip path.
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

export function baseName(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? path : path.slice(idx + 1);
}

// Path to the sibling .rels file for a given part.
export function relsPathFor(partPath: string): string {
  return `${dirOf(partPath)}/_rels/${baseName(partPath)}.rels`;
}

// r:id lookup tolerant of namespace-prefixed attributes across DOMParser impls.
export function relId(el: Element | null): string | null {
  if (!el) return null;
  return el.getAttributeNS(NS.rel, "id") ?? el.getAttribute("r:id");
}

// All direct element children of `parent`.
export function elementChildren(parent: Element): Element[] {
  const out: Element[] = [];
  for (const node of Array.from(parent.childNodes)) {
    if (node.nodeType === 1) out.push(node as Element);
  }
  return out;
}

// First DIRECT child matching ns+local. Use this (not firstChildNS) wherever
// structural position matters — e.g. a spPr's own solidFill vs a line's
// solidFill nested deeper, which a descendant search would wrongly grab.
export function directChild(parent: Element, ns: string, local: string): Element | null {
  for (const node of Array.from(parent.childNodes)) {
    if (node.nodeType === 1) {
      const el = node as Element;
      if (el.namespaceURI === ns && el.localName === local) return el;
    }
  }
  return null;
}

// All direct children matching ns+local.
export function childrenNS(parent: Element, ns: string, local: string): Element[] {
  return elementChildren(parent).filter(
    (e) => e.namespaceURI === ns && e.localName === local,
  );
}

// First DESCENDANT matching ns+local. Use only when the target is unambiguous
// under `parent` (e.g. locating clrScheme, spTree).
export function firstChildNS(parent: Element, ns: string, local: string): Element | null {
  const found = parent.getElementsByTagNameNS(ns, local);
  return found.length > 0 ? found[0] : null;
}

export function intAttr(el: Element | null, name: string, def = 0): number {
  if (!el) return def;
  const n = parseInt(el.getAttribute(name) ?? "", 10);
  return Number.isNaN(n) ? def : n;
}
