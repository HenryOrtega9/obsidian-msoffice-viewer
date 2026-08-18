import JSZip from "jszip";

// Count words in a .docx the way Word's status bar does (approximately): the
// main body text only, whitespace-delimited tokens. Paragraphs are separated
// so a run ending one paragraph never glues onto the start of the next.
// Headers, footers, footnotes, and text boxes outside the body flow are not
// counted, which matches Word's default "words in document" figure closely.
export async function countDocxWords(buf: ArrayBuffer): Promise<number | null> {
  try {
    const zip = await JSZip.loadAsync(buf);
    const doc = zip.file("word/document.xml");
    if (!doc) return null;
    const xml = await doc.async("string");
    const parser = new DOMParser();
    const dom = parser.parseFromString(xml, "application/xml");
    if (dom.getElementsByTagName("parsererror").length > 0) return null;
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const paras = dom.getElementsByTagNameNS(W, "p");
    let words = 0;
    for (let i = 0; i < paras.length; i++) {
      // Walk descendants in document order so a tab or line break between two
      // runs acts as a separator instead of gluing "foo<tab>bar" into one word.
      let text = "";
      const walker = dom.createTreeWalker(paras[i], NodeFilter.SHOW_ELEMENT);
      let node = walker.nextNode() as Element | null;
      while (node) {
        if (node.namespaceURI === W) {
          if (node.localName === "t") text += node.textContent ?? "";
          else if (node.localName === "tab" || node.localName === "br" || node.localName === "cr") text += " ";
        }
        node = walker.nextNode() as Element | null;
      }
      words += text.split(/\s+/).filter((t) => t.length > 0).length;
    }
    return words;
  } catch {
    return null;
  }
}
