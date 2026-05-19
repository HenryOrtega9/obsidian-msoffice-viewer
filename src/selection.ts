import type { StructuralLocator } from "./types";

const CONTEXT_CHARS = 200;

function findParagraphAncestor(
  node: Node,
  container: HTMLElement,
): HTMLElement | null {
  let n: Node | null = node;
  while (n && n !== container) {
    if (n.nodeType === Node.ELEMENT_NODE) {
      const el = n as HTMLElement;
      if (el.tagName === "P") return el;
    }
    n = n.parentNode;
  }
  return null;
}

function paragraphList(container: HTMLElement): HTMLElement[] {
  return Array.from(container.querySelectorAll("p"));
}

function offsetWithinParagraph(
  paragraph: HTMLElement,
  target: Node,
  targetOffset: number,
): number {
  let acc = 0;
  const walker = document.createTreeWalker(
    paragraph,
    NodeFilter.SHOW_TEXT,
    null,
  );
  let current: Node | null = walker.nextNode();
  while (current) {
    if (current === target) {
      return acc + targetOffset;
    }
    acc += (current.nodeValue ?? "").length;
    current = walker.nextNode();
  }
  return acc;
}

function inlineStyleSignature(node: Node): string {
  let n: Node | null = node;
  while (n && n.nodeType !== Node.ELEMENT_NODE) n = n.parentNode;
  if (!n) return "";
  const el = n as HTMLElement;
  return el.getAttribute("style") ?? "";
}

function spanCrossesFormatting(range: Range): boolean {
  const startSig = inlineStyleSignature(range.startContainer);
  const endSig = inlineStyleSignature(range.endContainer);
  return startSig !== endSig;
}

export function computeLocator(
  range: Range,
  containerEl: HTMLElement,
): StructuralLocator | null {
  const paragraph = findParagraphAncestor(range.startContainer, containerEl);
  if (!paragraph) return null;

  const endParagraph = findParagraphAncestor(range.endContainer, containerEl);
  if (endParagraph !== paragraph) {
    return {
      paragraphIndex: paragraphList(containerEl).indexOf(paragraph),
      startOffset: 0,
      endOffset: paragraph.textContent?.length ?? 0,
      selectedText: range.toString(),
      paragraphText: paragraph.textContent ?? "",
      surroundingContext: paragraphSurroundingContext(paragraph, containerEl),
      crossesFormatting: true,
    };
  }

  const paragraphs = paragraphList(containerEl);
  const paragraphIndex = paragraphs.indexOf(paragraph);
  const paragraphText = paragraph.textContent ?? "";
  const startOffset = offsetWithinParagraph(
    paragraph,
    range.startContainer,
    range.startOffset,
  );
  const endOffset = offsetWithinParagraph(
    paragraph,
    range.endContainer,
    range.endOffset,
  );
  const selectedText = paragraphText.slice(startOffset, endOffset);

  return {
    paragraphIndex,
    startOffset,
    endOffset,
    selectedText,
    paragraphText,
    surroundingContext: paragraphSurroundingContext(paragraph, containerEl),
    crossesFormatting: spanCrossesFormatting(range),
  };
}

function paragraphSurroundingContext(
  paragraph: HTMLElement,
  containerEl: HTMLElement,
): string {
  const paragraphs = paragraphList(containerEl);
  const idx = paragraphs.indexOf(paragraph);
  const before = paragraphs[idx - 1]?.textContent ?? "";
  const after = paragraphs[idx + 1]?.textContent ?? "";
  const beforeTrim = before.slice(-CONTEXT_CHARS);
  const afterTrim = after.slice(0, CONTEXT_CHARS);
  return `${beforeTrim}\n[...paragraph...]\n${afterTrim}`.trim();
}
