---
name: live-word-editing
description: Procedure for editing, formatting, or inspecting a .docx with the word-docx MCP tools when the file may be open in Microsoft Word (live mode via COM). Use whenever the user mentions a Word document being open, watching changes live, track changes, undo of agent edits, or asks to edit, restyle, or format a .docx with these tools, even if they do not say "live". Covers backend detection, safe mutation workflow, reading semantics for tracked deletions and tables, error triage, and the styling pipeline in live mode.
---

# Live Word Editing

Invariants (never close/save/reopen, use returned outline, stop on listed errors) are in the `live-word-mode` rule. This skill is the procedure.

## Workflow
1. Make the first tool call a read (`read_document`, `outline`). Note the backend prefix (`[backend: live]` or `[backend: headless]`). Why: it tells you whether edits appear in Word and whether undo is available.
2. Pick the target paragraph from the outline. Indices are 0-based and count visible body paragraphs only (tracked-deleted paragraphs and table-cell paragraphs are not `[p:N]`).
3. Mutate with `mutate_element` or `find_and_replace`. Read the returned local outline and take the next index from it. Why: indices shift after inserts and deletes.
4. After a batch of edits, tell the user the backend used, how many edits were made, and (if live) that each can be undone with one `Ctrl + Z`; content edits can also be reviewed as tracked changes.
5. Do not ask the user to close the file. Closing simply switches the next call to headless.

## Reading Semantics
- Reads show the final visible view: tracked-deleted text is filtered out and deleted paragraphs are skipped in numbering.
- A tracked change you made is still reversible by the user; do not "clean up" after yourself.

## Tracking Policy
- **Tracked:** `mutate_element`, `find_and_replace`.
- **Untracked (grouped in single UndoRecord):** `set_page_layout`, `update_styles`, `format_table_grid`, `compose_document`, `clean_structure`.

## Styling Pipeline in Live Mode
Follow `document-styling`. All 8 tools are fully live-capable:
1. `set_page_layout`: Sets page dimensions and margins across all sections in points.
2. `update_styles`: Updates font, bold, color, line spacing, and paragraph formatting on document styles.
3. `format_table_grid`: Sets table borders, cantSplit, repeat_header, and header shading.
4. `clean_structure`: Run with `dry_run=True` first, show the user the summary, then run with `dry_run=False`.

## When Something Goes Wrong
Stop per `problem-resolve`. Look up the message in `references/error-catalog.md` for what to tell the user and what not to do. Open that file only when a tool returns an error. If a call times out because Word may be blocked by a dialog, tell the user to dismiss it; do not retry.

## Known Limits
- Windows desktop Word only; attached to active Word session via COM (`GetActiveObject`).
- Replacing text takes on the formatting of the first replaced character.
- Automatic backups are saved to `documents/.backups/` before every mutating operation; the newest 20 per file stem are kept.

