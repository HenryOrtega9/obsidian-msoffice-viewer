// exceljs's built-in number-format table omits the implied currency formats
// (ids 5-8) and the accounting formats the Comma/Currency ribbon styles write
// (ids 41-44), so cells using them surface numFmt: undefined and render as
// bare General numbers. Its id 22 also quotes the hour ('m/d/yy "h":mm'),
// rendering a literal "h". Patch the shared table (consulted by
// StylesXform.getStyleModel) with the ECMA-376 §18.8.30 implied codes.

/* eslint-disable @typescript-eslint/no-var-requires */
// eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
const defaultNumFormats = require("exceljs/lib/xlsx/defaultnumformats") as Record<
  number,
  { f?: string }
>;

const MISSING_BUILTINS: Record<number, string> = {
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
};

let patched = false;

export function patchExcelJsDefaultNumFmts(): void {
  if (patched) return;
  patched = true;
  for (const [id, f] of Object.entries(MISSING_BUILTINS)) {
    const key = Number(id);
    if (!defaultNumFormats[key]) defaultNumFormats[key] = { f };
  }
  defaultNumFormats[22] = { f: "m/d/yy h:mm" };
}
