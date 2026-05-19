# Custom Obsidian Plugin: Preview .docx + Claude-Mediated Edits

## Context

You want to open `.docx` files inside Obsidian for **reading**, not as a fully editable WYSIWYG surface. When a change is needed, the workflow is:

1. Preview the doc inline in Obsidian
2. Select a span of text (or a section)
3. Invoke a command → tell Claude what to change
4. Claude rewrites that span
5. Plugin patches the underlying `.docx` and re-renders the preview

This trades direct manual editing for a Claude-as-editor pattern, which fits the broader strategy of routing edits through `claude-cli-chat` rather than building parallel editing surfaces. Saved as memory: [[feedback-claude-as-editor]].

### Why this is the right shape

- **Sidesteps the docx WYSIWYG fidelity cliff.** No JS library round-trips arbitrary docx with full fidelity. By making edits programmatic and targeted (rather than freeform keystroke-level), we only need to faithfully preserve the *parts we didn't touch* — which is a much easier problem than generic docx editing.
- **Reuses claude-cli-chat as the universal action layer.** The Claude bridge already exists, already routes to the Agent SDK credit pool, already handles auth. We don't rebuild it.
- **Read-only preview is a solved problem.** `docx-preview` renders docx as HTML/CSS in the browser with no edit complexity. Battle-tested.

---

## Architecture

### High-level flow

```
┌─────────────────────────────────────────────────────────────────┐
│  Obsidian main pane: DocxPreviewView                            │
│                                                                  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  docx-preview renders MyDoc.docx as HTML                 │   │
│  │  user selects: "The quarterly revenue grew 12%"          │   │
│  └──────────────────────────────────────────────────────────┘   │
│                          │                                       │
│                          │  ⌘E or palette: "Ask Claude to edit" │
│                          ▼                                       │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  Prompt modal: "rewrite this in past tense, add context  │   │
│  │  about Q3 headwinds"                                     │   │
│  └──────────────────────────────────────────────────────────┘   │
│                          │                                       │
└──────────────────────────┼───────────────────────────────────────┘
                           ▼
              app.plugins.plugins["claude-cli-chat"]
                           │   send: { selection, instruction, doc_outline }
                           ▼
                    claude-cli-chat → claude --print
                           │   structured response:
                           │   { type: "replace_text" | "rewrite_section",
                           │     target: <locator>, new_content: "..." }
                           ▼
                    EditDispatcher (this plugin)
                           │
            ┌──────────────┴──────────────┐
            ▼                             ▼
   OOXML in-place patch         Markdown round-trip
   (simple text edits           (structural edits:
    within one run group)        adds, deletes, reflow)
            │                             │
            └──────────────┬──────────────┘
                           ▼
                  Write MyDoc.docx back
                           │
                           ▼
                  Refresh docx-preview render
```

### File layout

```
obsidian-docx-claude/
  manifest.json
  package.json
  esbuild.config.mjs
  tsconfig.json
  styles.css
  src/
    main.ts                 # Plugin entrypoint, command registration
    DocxPreviewView.ts      # FileView — mounts docx-preview, captures selection
    ClaudeBridge.ts         # Talks to claude-cli-chat via app.plugins.plugins
    EditDispatcher.ts       # Routes Claude's response to the right edit path
    edits/
      ooxmlTextReplace.ts   # Path A: in-place text patching in word/document.xml
      markdownRoundtrip.ts  # Path B: extract → edit → convert back
    selection.ts            # Maps DOM selection back to docx structural locator
    settings.ts
    types.ts
```

### Key dependencies

- `docx-preview` — read-only render of docx to HTML
- `jszip` — read/write the docx archive (docx is a zip of XML files)
- `fast-xml-parser` or `@xmldom/xmldom` — parse/serialize `word/document.xml`
- `mammoth` — docx → markdown for the round-trip path
- `pandoc` (external binary, shelled out) — markdown → docx for the round-trip path. Plugin checks for `pandoc` on load and surfaces a clear install instruction if missing. (Alternative: `md-to-docx` pure-JS, but pandoc gives better fidelity and you already have it in the Pandoc-plugin ecosystem.)

### Module responsibilities

**`DocxPreviewView.ts`** — extends `FileView`, declares view type `"docx-claude-view"`. On `onLoadFile`:
- `readBinary` the .docx → ArrayBuffer
- Pass to `docx-preview` (`renderAsync(buffer, containerEl)`)
- Attach a `mouseup`/`selectionchange` listener to track the user's selection within the rendered HTML
- Stash the current selection (text + DOM range) in view state

**`selection.ts`** — given a DOM `Range` in the docx-preview output, compute a *structural locator* that survives the OOXML edit:
- Walk up to the nearest paragraph element (`<p>` with a docx-preview-assigned id, or compute paragraph index)
- Record: `{ paragraphIndex, startOffset, endOffset, selectedText, surroundingContext }`
- This locator is what gets sent to Claude *and* what `EditDispatcher` uses to find the right run in `word/document.xml`

**`ClaudeBridge.ts`** — looks up `app.plugins.plugins["claude-cli-chat"]` at runtime. Verifies the plugin is enabled. Calls its exposed method (verify the exact API surface at implementation time — read the plugin's `main.ts` to see what's on the public interface). Falls back to a clear error toast if claude-cli-chat is missing or doesn't expose the needed entrypoint. Sends Claude a system prompt instructing it to return structured JSON:

```json
{
  "type": "replace_text" | "rewrite_section" | "insert_after" | "delete",
  "target_locator": { "paragraphIndex": 12, ... },
  "new_content": "...",
  "fidelity_hint": "text_only" | "structural"
}
```

**`EditDispatcher.ts`** — switches on `type` + `fidelity_hint`:
- `replace_text` + `text_only` → `ooxmlTextReplace`
- `rewrite_section` / `insert_after` / structural → `markdownRoundtrip`
- Always writes a one-shot `.bak` of the original on first edit per session

**`edits/ooxmlTextReplace.ts`** — the careful path:
- Unzip with jszip, parse `word/document.xml`
- Walk to the target paragraph by index
- Concatenate visible text across child `<w:r>` runs to build a flat string
- Locate the selection offsets in that flat string
- Modify the run sequence: split runs at offset boundaries if needed, swap text in the middle runs, preserve the run-property (`<w:rPr>`) of the original surrounding runs so bold/italic/font/color is kept
- Serialize, re-zip, write back
- **Hard case to handle:** selection spans runs with different formatting (e.g. "Hello **world**" where user changes "Hello world" → "Goodbye world"). Strategy: if the selection crosses run-property boundaries, escalate to the markdown round-trip path automatically. Don't try to be clever about merging formatting.

**`edits/markdownRoundtrip.ts`** — the fallback path:
- mammoth converts the whole docx → markdown
- Apply Claude's edit to the markdown string (Claude returns the *new markdown* for the section)
- Shell out: `pandoc -f markdown -t docx --reference-doc=<original.docx> -o <out.docx>` — the `--reference-doc` flag is crucial; it preserves the original's styles (fonts, margins, heading definitions) so the output looks like a continuation of the source doc, not a fresh pandoc default
- Replace the file
- Acknowledged fidelity loss: anything mammoth can't represent in markdown (track changes, comments, complex layout) is gone after a round-trip. Document this clearly in the README and plugin settings.

**`main.ts`** — registers:
- View type `docx-claude-view`
- Extensions `["docx"]` → that view
- Command "Docx: Ask Claude to edit selection" (hotkey-bindable)
- Command "Docx: Convert to markdown note" (one-shot bypass via mammoth, for users who want to defect to markdown permanently)
- Settings tab

### Integration with claude-cli-chat

**Action required at implementation time:** open `<vault>/.obsidian/plugins/claude-cli-chat/` and read its `main.ts` (or built `main.js`) to confirm what methods it exposes on its plugin instance. Likely candidates:
- A `sendPrompt(text, opts)` method
- An event bus the plugin emits on
- A workspace command we can trigger via `app.commands.executeCommandById`

If claude-cli-chat doesn't currently expose a programmatic API, the small addition there is: add an exported method like `async runHeadlessPrompt(systemPrompt: string, userPrompt: string): Promise<string>` that calls the same `claude --print` path it uses internally and returns the stdout. That's a one-function addition to a plugin you already maintain.

---

## Critical files

- `src/main.ts` — registration glue
- `src/DocxPreviewView.ts` — the user-facing view (selection capture happens here)
- `src/ClaudeBridge.ts` — the integration point with claude-cli-chat, single source of truth for the prompt format
- `src/edits/ooxmlTextReplace.ts` — the high-fidelity edit path, hardest code in the project
- `src/edits/markdownRoundtrip.ts` — the safety-net edit path
- **External (not modified initially, but to verify):** `<vault>/.obsidian/plugins/claude-cli-chat/main.ts` to confirm the bridge surface

---

## Known risks

1. **claude-cli-chat may not expose a public API today.** Mitigation: read its source first; if needed, add a small `runHeadlessPrompt` exported method there before building this plugin.
2. **OOXML run-splitting is genuinely fiddly.** Mitigation: escalate to markdown round-trip when the selection crosses formatting boundaries — don't try to perfect the OOXML path.
3. **`pandoc` requirement.** Mitigation: detect on plugin load, show actionable install instructions (`brew install pandoc`).
4. **docx-preview rendering doesn't 1:1 match Word's layout.** That's fine for read + edit-via-Claude; the user isn't trying to use this as a layout authoring tool. Note it in the README.
5. **Claude returning malformed structured JSON.** Mitigation: schema-validate the response in `ClaudeBridge`; on parse failure, fall back to dumping Claude's raw text as the new markdown for the targeted section and routing through the round-trip path.
6. **Backup hygiene.** Per-session `.bak` is cheap; for paranoia, also keep last N edits in `<vault>/.obsidian/plugins/obsidian-docx-claude/history/<filename>/<timestamp>.docx`.

---

## Verification

1. **Install + open.** Build plugin, copy to `<vault>/.obsidian/plugins/obsidian-docx-claude/`, enable. Drop a simple `.docx` in the vault, click it — confirm preview renders.
2. **Selection capture.** Select a phrase, run "Docx: Ask Claude to edit selection". Confirm the prompt modal opens with the selection echoed back.
3. **Simple text edit (OOXML path).** Select an unformatted sentence, ask Claude to rephrase. Confirm: file updates, preview refreshes, surrounding paragraphs and styling unchanged. Open the saved .docx in Word/Pages and confirm formatting integrity.
4. **Cross-formatting edit (round-trip path).** Select a sentence that spans bold/italic boundaries. Confirm the dispatcher routes to the markdown round-trip path and the result is still openable in Word.
5. **Structural edit.** Ask Claude to "insert a new paragraph after this one summarizing the prior section." Confirm the new paragraph appears, original content is preserved.
6. **Round-trip stress test.** Apply 5 consecutive edits to the same doc. Confirm progressive fidelity loss is acceptable (or surface a "drift warning" after N round-trip edits in the same session).
7. **Backup files.** Confirm `.bak` is created on first edit and `history/` snapshots accumulate per setting.
8. **claude-cli-chat absent.** Disable claude-cli-chat, try to invoke an edit — confirm graceful error with install/enable instructions.
9. **Mobile.** Set `isDesktopOnly: true` in `manifest.json` until pandoc-on-mobile is sorted (it isn't).
