import JSZip from "jszip";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { EscalateToRoundtrip } from "../types";
import type { StructuralLocator } from "../types";

/*
 * Path A: in-place text patching of word/document.xml.
 *
 * Strategy:
 *  - Parse with preserveOrder:true so we round-trip structure verbatim.
 *  - Walk to the Nth <w:p> in document order.
 *  - Flatten its <w:r>/<w:t> children into a plain string with index map.
 *  - Locate the selection's [start, end) range in that flat string.
 *  - If the range sits inside a single run, replace that run's text.
 *  - If it spans multiple runs sharing identical w:rPr, merge them into one
 *    run with the new text.
 *  - Otherwise: escalate to round-trip - we do not attempt to be clever about
 *    crossing formatting boundaries (per PLAN.md).
 */

type OrderedNode = Record<string, unknown> & { ":@"?: Record<string, unknown> };

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
});

const builder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  format: false,
  suppressEmptyNode: false,
});

function isTag(node: OrderedNode, tag: string): boolean {
  return Object.prototype.hasOwnProperty.call(node, tag);
}

function findBody(root: OrderedNode[]): OrderedNode[] | null {
  for (const top of root) {
    if (isTag(top, "w:document")) {
      const docKids = top["w:document"] as OrderedNode[];
      for (const c of docKids) {
        if (isTag(c, "w:body")) {
          return c["w:body"] as OrderedNode[];
        }
      }
    }
  }
  return null;
}

function getParagraphChildren(p: OrderedNode): OrderedNode[] {
  return (p["w:p"] as OrderedNode[]) ?? [];
}

function getRunChildren(r: OrderedNode): OrderedNode[] {
  return (r["w:r"] as OrderedNode[]) ?? [];
}

interface RunInfo {
  node: OrderedNode;
  text: string;
  start: number;
  end: number;
  rPr: string;
}

function extractRunText(run: OrderedNode): string {
  let acc = "";
  for (const child of getRunChildren(run)) {
    if (isTag(child, "w:t")) {
      const tKids = child["w:t"] as OrderedNode[];
      for (const tk of tKids) {
        const txt = (tk as { "#text"?: string })["#text"];
        if (typeof txt === "string") acc += txt;
      }
    } else if (isTag(child, "w:tab")) {
      acc += "\t";
    } else if (isTag(child, "w:br")) {
      acc += "\n";
    }
  }
  return acc;
}

function rPrSignature(run: OrderedNode): string {
  for (const child of getRunChildren(run)) {
    if (isTag(child, "w:rPr")) {
      return JSON.stringify(child["w:rPr"]);
    }
  }
  return "";
}

function setRunText(run: OrderedNode, newText: string): void {
  const kids = getRunChildren(run);
  let firstTextIdx = -1;
  for (let i = 0; i < kids.length; i++) {
    if (isTag(kids[i], "w:t")) {
      firstTextIdx = i;
      break;
    }
  }
  const newTextNode: OrderedNode = {
    "w:t": [{ "#text": newText } as OrderedNode],
    ":@": { "@_xml:space": "preserve" },
  };
  if (firstTextIdx === -1) {
    kids.push(newTextNode);
  } else {
    const replacement: OrderedNode[] = [];
    for (let i = 0; i < kids.length; i++) {
      const k = kids[i];
      if (i === firstTextIdx) {
        replacement.push(newTextNode);
        continue;
      }
      if (isTag(k, "w:t") || isTag(k, "w:tab") || isTag(k, "w:br")) continue;
      replacement.push(k);
    }
    run["w:r"] = replacement;
  }
}

function buildRunInfos(paragraph: OrderedNode): {
  runs: RunInfo[];
  flat: string;
} {
  const runs: RunInfo[] = [];
  let cursor = 0;
  let flat = "";
  for (const child of getParagraphChildren(paragraph)) {
    if (!isTag(child, "w:r")) continue;
    const text = extractRunText(child);
    const start = cursor;
    const end = cursor + text.length;
    runs.push({ node: child, text, start, end, rPr: rPrSignature(child) });
    flat += text;
    cursor = end;
  }
  return { runs, flat };
}

function findParagraphByIndex(
  body: OrderedNode[],
  index: number,
): OrderedNode | null {
  let count = 0;
  for (const node of body) {
    if (isTag(node, "w:p")) {
      if (count === index) return node;
      count++;
    }
  }
  return null;
}

export async function ooxmlTextReplace(
  docxBuffer: ArrayBuffer,
  locator: StructuralLocator,
  newText: string,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(docxBuffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) {
    throw new EscalateToRoundtrip("word/document.xml not found in archive");
  }
  const xml = await docFile.async("string");
  const parsed = parser.parse(xml) as OrderedNode[];
  const body = findBody(parsed);
  if (!body) {
    throw new EscalateToRoundtrip("document body not found");
  }
  const paragraph = findParagraphByIndex(body, locator.paragraphIndex);
  if (!paragraph) {
    throw new EscalateToRoundtrip(
      `paragraph #${locator.paragraphIndex} not found`,
    );
  }

  const { runs, flat } = buildRunInfos(paragraph);
  const { startOffset, endOffset } = locator;

  if (
    startOffset < 0 ||
    endOffset > flat.length ||
    startOffset > endOffset
  ) {
    throw new EscalateToRoundtrip("locator offsets out of range");
  }

  const touched = runs.filter(
    (r) => !(r.end <= startOffset) && !(r.start >= endOffset),
  );
  if (touched.length === 0) {
    throw new EscalateToRoundtrip("no runs intersect the selection");
  }

  const firstSig = touched[0].rPr;
  const sameFormatting = touched.every((r) => r.rPr === firstSig);
  if (!sameFormatting) {
    throw new EscalateToRoundtrip("selection spans differing run-property sets");
  }

  const replaced = flat.slice(0, startOffset) + newText + flat.slice(endOffset);
  const firstRun = touched[0];
  const lastRun = touched[touched.length - 1];
  const delta = newText.length - (endOffset - startOffset);
  const sliceEnd = lastRun.end + delta;
  const newRunText = replaced.slice(firstRun.start, sliceEnd);
  setRunText(firstRun.node, newRunText);

  if (touched.length > 1) {
    const toRemove = new Set(touched.slice(1).map((r) => r.node));
    paragraph["w:p"] = getParagraphChildren(paragraph).filter(
      (c) => !toRemove.has(c),
    );
  }

  const newXml = builder.build(parsed);
  zip.file("word/document.xml", newXml);
  const out = await zip.generateAsync({ type: "uint8array" });
  return out;
}
