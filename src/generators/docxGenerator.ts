import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from "docx";
import { DocxBlock, DocxSpec } from "./spec";

function headingLevel(level: 1 | 2 | 3): (typeof HeadingLevel)[keyof typeof HeadingLevel] {
  if (level === 1) return HeadingLevel.HEADING_1;
  if (level === 2) return HeadingLevel.HEADING_2;
  return HeadingLevel.HEADING_3;
}

function textRunsFromCell(cell: string): TextRun[] {
  // Preserve cell-internal newlines in table cells (TextRun + break:1) so
  // multi-line cells render correctly in Word.
  const lines = cell.split("\n");
  const runs: TextRun[] = [];
  lines.forEach((line, idx) => {
    if (idx > 0) runs.push(new TextRun({ text: "", break: 1 }));
    runs.push(new TextRun(line));
  });
  return runs;
}

function blockToChildren(block: DocxBlock, numberingRef: string): (Paragraph | Table)[] {
  switch (block.type) {
    case "heading":
      return [
        new Paragraph({
          heading: headingLevel(block.level),
          children: [new TextRun(block.text)],
        }),
      ];
    case "paragraph":
      return [new Paragraph({ children: [new TextRun(block.text)] })];
    case "bullets":
      return block.items.map(
        (item) =>
          new Paragraph({
            text: item,
            bullet: { level: 0 },
          }),
      );
    case "numbered":
      return block.items.map(
        (item) =>
          new Paragraph({
            text: item,
            numbering: { reference: numberingRef, level: 0 },
          }),
      );
    case "table": {
      const rows = block.rows.map(
        (row) =>
          new TableRow({
            children: row.map(
              (cell) =>
                new TableCell({
                  children: [new Paragraph({ children: textRunsFromCell(cell) })],
                }),
            ),
          }),
      );
      return [new Table({ rows })];
    }
  }
}

export async function buildDocx(spec: DocxSpec): Promise<ArrayBuffer> {
  const children: (Paragraph | Table)[] = [];
  if (spec.title) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.TITLE,
        children: [new TextRun(spec.title)],
      }),
    );
  }
  // Each numbered block gets its own numbering definition so list counters
  // restart at 1. A single shared reference makes the second list continue
  // from the first.
  type NumberingConfigItem = NonNullable<
    ConstructorParameters<typeof Document>[0]["numbering"]
  >["config"][number];
  const numberingConfig: NumberingConfigItem[] = [];
  let numberedIdx = 0;
  spec.blocks.forEach((block) => {
    let ref = "";
    if (block.type === "numbered") {
      ref = `msoffice-numbered-${numberedIdx++}`;
      numberingConfig.push({
        reference: ref,
        levels: [
          { level: 0, format: "decimal", text: "%1.", alignment: "start" },
        ],
      });
    }
    for (const child of blockToChildren(block, ref)) children.push(child);
  });

  const doc = new Document({
    numbering: { config: numberingConfig },
    sections: [{ children }],
  });

  const blob = await Packer.toBlob(doc);
  return await blob.arrayBuffer();
}
