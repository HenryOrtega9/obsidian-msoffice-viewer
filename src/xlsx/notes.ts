import type ExcelJS from "exceljs";
import { renderRichTextRuns, RichTextRun } from "./richText";

export interface CellNote {
  plain: string;
  runs: RichTextRun[] | null;
}

export function extractCellNote(cell: ExcelJS.Cell): CellNote | null {
  const note = cell.note as unknown;
  if (note == null || note === "") return null;
  if (typeof note === "string") return { plain: note, runs: null };
  if (typeof note === "object") {
    const obj = note as { texts?: RichTextRun[] };
    if (Array.isArray(obj.texts)) {
      const plain = obj.texts.map((t) => t.text ?? "").join("");
      return { plain, runs: obj.texts };
    }
  }
  return null;
}

// Attach a red corner triangle to the cell that:
//   - on hover, surfaces the note as a native title tooltip
//   - on click, toggles a popover anchored to the cell with the full text
// The popover is appended to the grid container (passed in as `popoverHost`)
// so it can escape any cell-level overflow:hidden.
export function addNoteMarker(
  td: HTMLTableCellElement,
  note: CellNote,
  popoverHost: HTMLElement,
): void {
  td.title = note.plain;
  // Cells default to position: static; sticky frozen cells become "sticky".
  // The marker needs an explicit positioning context regardless. We piggyback
  // on inline style so we don't fight sticky positioning when applied later.
  if (!td.style.position) td.style.position = "relative";

  const marker = td.createSpan({ cls: "docx-claude-xlsx-note-marker" });
  marker.addEventListener("click", (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    toggleNotePopover(td, note, popoverHost);
  });
}

let activePopover: HTMLElement | null = null;
let activeOutsideHandler: ((ev: MouseEvent) => void) | null = null;

function toggleNotePopover(anchor: HTMLElement, note: CellNote, host: HTMLElement): void {
  if (activePopover) {
    closeActivePopover();
    return;
  }
  const popover = host.createDiv({ cls: "docx-claude-xlsx-note-popover" });
  if (note.runs) renderRichTextRuns(note.runs, popover);
  else popover.setText(note.plain);

  // Position to the right of the cell, clamped to host width.
  const hostRect = host.getBoundingClientRect();
  const cellRect = anchor.getBoundingClientRect();
  const left = Math.min(
    cellRect.right - hostRect.left + host.scrollLeft,
    host.scrollWidth - 240,
  );
  const top = cellRect.top - hostRect.top + host.scrollTop;
  popover.style.left = `${Math.max(0, left)}px`;
  popover.style.top = `${Math.max(0, top)}px`;

  activePopover = popover;
  activeOutsideHandler = (ev: MouseEvent) => {
    if (popover.contains(ev.target as Node)) return;
    closeActivePopover();
  };
  // Defer so the click that opened us doesn't immediately close it.
  setTimeout(() => {
    if (activeOutsideHandler) document.addEventListener("click", activeOutsideHandler);
  }, 0);
}

// Exported so the view can tear down a popover (and its document-level
// outside-click listener) on sheet switch, file switch, and view unload —
// otherwise the listener and module globals leak across renders.
export function closeActivePopover(): void {
  if (activePopover) {
    activePopover.remove();
    activePopover = null;
  }
  if (activeOutsideHandler) {
    document.removeEventListener("click", activeOutsideHandler);
    activeOutsideHandler = null;
  }
}
