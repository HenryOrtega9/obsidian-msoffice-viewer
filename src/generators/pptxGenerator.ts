import PptxGenJS from "pptxgenjs";
import { PptxSpec } from "./spec";

const SLIDE_WIDTH_IN = 13.333;
const SLIDE_HEIGHT_IN = 7.5;
const CUSTOM_LAYOUT = "MSOFFICE_VIEWER_WIDE";

export async function buildPptx(spec: PptxSpec): Promise<ArrayBuffer> {
  const pres = new PptxGenJS();
  // defineLayout + assign by name is portable across pptxgenjs versions; the
  // string "LAYOUT_WIDE" assignment is brittle.
  pres.defineLayout({
    name: CUSTOM_LAYOUT,
    width: SLIDE_WIDTH_IN,
    height: SLIDE_HEIGHT_IN,
  });
  pres.layout = CUSTOM_LAYOUT;

  // No auto-title-slide here: the system prompt already instructs Claude to
  // produce a title slide as the first entry of `slides` when appropriate.
  // Inserting one here on top produced duplicate title slides.

  for (const slide of spec.slides) {
    const s = pres.addSlide();
    let y = 0.5;

    if (slide.title) {
      s.addText(slide.title, {
        x: 0.5,
        y,
        w: SLIDE_WIDTH_IN - 1,
        h: 0.9,
        fontSize: 32,
        bold: true,
      });
      y += 1.1;
    }

    if (slide.bullets && slide.bullets.length > 0) {
      s.addText(
        slide.bullets.map((b) => ({ text: b, options: { bullet: true } })),
        {
          x: 0.7,
          y,
          w: SLIDE_WIDTH_IN - 1.4,
          h: SLIDE_HEIGHT_IN - y - 0.5,
          fontSize: 20,
          valign: "top",
        },
      );
    } else if (slide.body) {
      s.addText(slide.body, {
        x: 0.7,
        y,
        w: SLIDE_WIDTH_IN - 1.4,
        h: SLIDE_HEIGHT_IN - y - 0.5,
        fontSize: 20,
        valign: "top",
      });
    }

    if (slide.notes) {
      // Strip non-printable / XML-illegal control chars so PowerPoint can
      // open the file. Leave \t \n \r alone.
      s.addNotes(slide.notes.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ""));
    }
  }

  const out = (await pres.write({ outputType: "arraybuffer" })) as ArrayBuffer | Uint8Array;
  if (out instanceof ArrayBuffer) return out;
  const u8 = out as Uint8Array;
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
}
