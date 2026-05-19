// Standalone smoke test for ooxmlTextReplace.
// Run: npx tsx src/__test__/ooxml.test.mts
import JSZip from "jszip";
import { ooxmlTextReplace } from "../edits/ooxmlTextReplace.js";

function makeDocx(paragraphs: Array<Array<{ text: string; bold?: boolean }>>): Promise<Uint8Array> {
  const body = paragraphs
    .map((runs) => {
      const rs = runs
        .map((r) => {
          const rPr = r.bold ? "<w:rPr><w:b/></w:rPr>" : "";
          return `<w:r>${rPr}<w:t xml:space="preserve">${r.text}</w:t></w:r>`;
        })
        .join("");
      return `<w:p>${rs}</w:p>`;
    })
    .join("");

  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${body}</w:body>
</w:document>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rels);
  zip.file("word/document.xml", doc);
  return zip.generateAsync({ type: "uint8array" });
}

async function extractParagraphTexts(buffer: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");
  const paragraphs: string[] = [];
  const pMatches = xml.matchAll(/<w:p[^>]*>([\s\S]*?)<\/w:p>/g);
  for (const pm of pMatches) {
    const inner = pm[1];
    const texts = Array.from(inner.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)).map(
      (m) => m[1],
    );
    paragraphs.push(texts.join(""));
  }
  return paragraphs;
}

async function run() {
  {
    const buf = await makeDocx([[{ text: "Hello world" }]]);
    const out = await ooxmlTextReplace(
      buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      {
        paragraphIndex: 0,
        startOffset: 0,
        endOffset: 5,
        selectedText: "Hello",
        paragraphText: "Hello world",
        surroundingContext: "",
        crossesFormatting: false,
      },
      "Goodbye",
    );
    const ps = await extractParagraphTexts(out);
    if (ps[0] !== "Goodbye world") throw new Error(`case 1: got ${JSON.stringify(ps)}`);
    console.log("case 1 ok:", ps[0]);
  }

  {
    const buf = await makeDocx([
      [{ text: "Hello " }, { text: "world" }, { text: "!" }],
    ]);
    const out = await ooxmlTextReplace(
      buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      {
        paragraphIndex: 0,
        startOffset: 0,
        endOffset: 12,
        selectedText: "Hello world!",
        paragraphText: "Hello world!",
        surroundingContext: "",
        crossesFormatting: false,
      },
      "Replaced",
    );
    const ps = await extractParagraphTexts(out);
    if (ps[0] !== "Replaced") throw new Error(`case 2: got ${JSON.stringify(ps)}`);
    console.log("case 2 ok:", ps[0]);
  }

  {
    const buf = await makeDocx([
      [{ text: "Hello " }, { text: "world", bold: true }],
    ]);
    let escalated = false;
    try {
      await ooxmlTextReplace(
        buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
        {
          paragraphIndex: 0,
          startOffset: 0,
          endOffset: 11,
          selectedText: "Hello world",
          paragraphText: "Hello world",
          surroundingContext: "",
          crossesFormatting: false,
        },
        "Goodbye world",
      );
    } catch (e: any) {
      if (e.name === "EscalateToRoundtrip") {
        escalated = true;
      } else {
        throw e;
      }
    }
    if (!escalated) throw new Error("case 3: expected escalation");
    console.log("case 3 ok: escalated as expected");
  }

  {
    const buf = await makeDocx([
      [{ text: "First" }],
      [{ text: "Second paragraph here" }],
    ]);
    const out = await ooxmlTextReplace(
      buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      {
        paragraphIndex: 1,
        startOffset: 7,
        endOffset: 16,
        selectedText: "paragraph",
        paragraphText: "Second paragraph here",
        surroundingContext: "",
        crossesFormatting: false,
      },
      "PARAGRAPH",
    );
    const ps = await extractParagraphTexts(out);
    if (ps[0] !== "First") throw new Error(`case 4 p0: got ${ps[0]}`);
    if (ps[1] !== "Second PARAGRAPH here") throw new Error(`case 4 p1: got ${ps[1]}`);
    console.log("case 4 ok:", ps);
  }

  console.log("\nAll cases passed.");
}

run().catch((e) => {
  console.error(e);
  process.exit(1);
});
