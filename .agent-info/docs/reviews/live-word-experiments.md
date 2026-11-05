# Phase 1 Experiments & Spikes Review

**Date:** 2026-11-06  
**Branch:** `feat/live-word-backend`  
**Word Version:** Microsoft Word for Microsoft 365 (Desktop Win32 COM)

---

## Summary of Experiments

| Experiment | Focus | Result | Baseline / Conclusion |
|---|---|---|---|
| **E1** | File reading while open in Word | **Verified** | Standard open files have `FILE_SHARE_READ`. `shutil.copy2` creates a valid standalone snapshot without crashing or throwing sharing violations. |
| **E2** | Lock file naming pattern | **Verified** | Lock files follow Word convention: `~$<filename>`. When filename is `live_test.docx`, lock is `~$live_test.docx`. Detecting `~$` accurately indicates active Word locks. |
| **E3** | Undo stack single-step (`UndoRecord`) | **Verified** | Wrapped operations with `app.UndoRecord.StartCustomRecord(...)` and `EndCustomRecord()` collapse all child mutations into exactly one entry in the Word Undo stack (single Ctrl+Z). |
| **E5** | Track changes revision behavior | **Verified** | Deleted paragraphs remain in `doc.Paragraphs` with `rng.Revisions(1).Type == wdRevisionDelete (2)`. Bulk range mapping via `doc.Tables` and body paragraph filtering prevents deleted/table paragraphs from shifting visible body indices. |
| **E6** | Performance benchmark (1,000 paragraphs) | **Passed** | 1,000 paragraphs + 10 headings benchmark executed in **0.433s** (budget: < 2.0s). Using `GetCrossReferenceItems` and bulk table range sets eliminates COM IPC round-trip overhead. |
| **E7** | Revision authoring | **Verified** | Word 365 connected accounts override `Application.UserName` with the signed-in account name. Edits will share the user's author identity seamlessly per Section 3.10. |

---

## Detailed Notes

### E3: Undo Record
```python
undo_record = app.UndoRecord
undo_record.StartCustomRecord(f"Agent: {tool_name} {target_desc}")
try:
    ...
finally:
    undo_record.EndCustomRecord()
```
Testing confirmed that even if a tool performs multiple paragraph insertions, deletions, and formatting operations, Microsoft Word groups them into a single named entry in the Undo dropdown. Pressing Ctrl+Z once cleanly reverts all changes made by the tool.

### E5 & E6: Paragraph Mapping & Outline Performance
- Iterating through 1,000 paragraphs one-by-one with `rng.Information(wdWithInTable)` required 3,000+ COM round-trips taking > 20s.
- Optimized design:
  1. Iterate `doc.Tables` (typically < 10) to map paragraph index intervals inside tables: `doc.Range(0, tbl.Range.Start).Paragraphs.Count + 1` to `+ tbl.Range.Paragraphs.Count`.
  2. Compute visible body indices as `[i for i in range(1, total + 1) if i not in table_paragraphs]`.
  3. Outline extraction queries `doc.GetCrossReferenceItems(1)` (wdRefTypeHeading) and searches matches forward with `rng.Find`, resolving paragraph indices via `doc.Range(0, rng.Paragraphs.Item(1).Range.End).Paragraphs.Count`.
  4. Execution time dropped from > 20s to **0.433s**, exceeding the < 2.0s requirement.

