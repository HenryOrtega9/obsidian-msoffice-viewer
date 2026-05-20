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

function blockToChildren(block: DocxBlock): (Paragraph | Table)[] {
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
            numbering: { reference: "msoffice-numbered", level: 0 },
          }),
      );
    case "table": {
      const rows = block.rows.map(
        (row) =>
          new TableRow({
            children: row.map(
              (cell) =>
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun(cell)] })],
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
  for (const block of spec.blocks) {
    for (const child of blockToChildren(block)) children.push(child);
  }

  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "msoffice-numbered",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: "start",
            },
          ],
        },
      ],
    },
    sections: [{ children }],
  });

  const blob = await Packer.toBlob(doc);
  return await blob.arrayBuffer();
}
