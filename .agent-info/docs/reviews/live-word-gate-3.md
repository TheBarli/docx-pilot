# Gate 3 Review: Hardening, Backups, and Safety

**Date:** 2026-11-06  
**Branch:** `feat/live-word-backend`  
**Status:** **PASSED**

---

## 1. Hardening Capabilities Implemented

### 1. Pre-mutation Automated Backups
- Added `create_backup(file_path)` hook executed before every mutating tool call in `src/server.py`.
- Tested and verified against both open live Word documents and on-disk files.
- Automatically creates timestamped copies in `documents/.backups/` (`{stem}_{YYYYMMDD_HHMMSS_micros}.docx`).
- Automatically rotates and retains the latest `BACKUP_KEEP = 20` backups per file stem.

### 2. Headers & Footers Scope in Find and Replace
- Extended `LiveWordEngine.find_and_replace` to support `scope in ("all", "body", "tables", "headers_footers")`.
- Searches across `doc.Content` as well as all section headers and footers (`sec.Headers` and `sec.Footers`).
- Retains case-matching, tracked deletion filtering, and 255 character limit guard.

### 3. Read-Only and Protected Document Safeguards
- `pick_backend` actively guards against mutating documents in protected view or read-only mode, surfacing user-friendly error messages advising the user to enable editing in Word.

### 4. Single-Step Undo (`UndoRecord`) Across All Mutations
- Every mutating tool runs under `live_call(...)` context manager wrapping operations inside a custom `UndoRecord`.
- Verified across insertions, replacements, deletions, styling updates, table formatting, and clean operations: pressing `Ctrl + Z` once cleanly undoes the full tool invocation.

### 5. Multi-Window Isolation
- Fixtures and tests isolate temporary test files inside pytest's `tmp_path`, preventing any lock conflicts with active documents open in Microsoft Word.

---

## 2. Test Suite Confirmation
- Full test suite verified green across headless and live test runs.
- Backups verified functioning in `documents/.backups/`.

