import PptxGenJS from "pptxgenjs";
import { PptxSpec } from "./spec";

const SLIDE_WIDTH_IN = 13.333;
const SLIDE_HEIGHT_IN = 7.5;

export async function buildPptx(spec: PptxSpec): Promise<ArrayBuffer> {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE";

  if (spec.title) {
    const titleSlide = pres.addSlide();
    titleSlide.addText(spec.title, {
      x: 0.5,
      y: SLIDE_HEIGHT_IN / 2 - 0.75,
      w: SLIDE_WIDTH_IN - 1,
      h: 1.5,
      fontSize: 44,
      bold: true,
      align: "center",
      valign: "middle",
    });
  }

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

    if (slide.notes) s.addNotes(slide.notes);
  }

  const out = (await pres.write({ outputType: "arraybuffer" })) as ArrayBuffer;
  return out;
}
