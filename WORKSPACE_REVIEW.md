# Workspace Review: MsWord-ai-integration

**Repository:** `TheBarli/MsWord-ai-integration`  
**Current Branch:** `feat/live-word-backend`  
**Base Tracking:** `origin/refactor/mcp-and-directory-cleanup` (Default: `origin/master`)  
**Environment:** Python 3.11.9 (`.venv`), FastMCP 4.0+, Windows COM (`pywin32` 312)  
**Status:** **Phases 0–2 and 4 implemented; Phase 3 partial (see Known Limitations) | Automated Gates 0, 2, 3 Passed | 59 Passed with Word Running (17 Skipped When Word Closed) | Pending Human Gates 1 & 4**

---

## 1. Executive Summary

This workspace implements an isolated, token-efficient **Model Context Protocol (MCP)** server for Microsoft Word (`.docx`) manipulation tailored for AI agent pair-programming (Antigravity IDE). It features a **dual-backend architecture**:
1. **Headless Backend (`DocxEngine`):** Fast, schema-safe `python-docx` + `lxml` manipulations with strict ECMA-376 XML child element ordering. Operates without requiring Microsoft Word installed or running.
2. **Live Backend (`LiveWordEngine`):** Direct Windows COM bridge to an active Microsoft Word desktop session. Enables real-time document viewing, one-click Ctrl+Z single-step undo records, and Track Changes for collaborative review.

All 8 tools across the unified MCP surface are implemented for both headless and live modes. Automated verification gates (Gate 0, Gate 2, Gate 3) have passed. Interactive verification gates (Gate 1 and Gate 4) are cataloged for human sign-off before merging.

---

## 2. Directory Architecture & Component Inventory

```text
MsWord-ai-integration/
├── .agents/                                # Agent configuration & formatting orchestration
│   ├── mcp_config.json                     # Portable workspace MCP registration (python src/server.py)
│   ├── presets/                            # Standardized YAML document formatting profiles
│   │   ├── academic-id.yaml                # Standard Indonesian academic format (A4, 4-4-3-3 cm)
│   │   ├── academic-id-advanced.yaml       # Fine-grained academic preset (H1-H5 hierarchy, spacing)
│   │   └── clean-uniform.yaml              # General business clean styling profile
│   ├── rules/                              # Agent behavioral and pipeline execution rules
│   │   ├── document-styling.md             # 5-step automated formatting orchestration rule
│   │   ├── live-word-mode.md               # Live Word mode behavioral constraints
│   │   └── [local dev rules]               # gh-usage, git-usage, etc. (excluded via .git/info/exclude)
│   └── skills/
│       └── live-word-editing/              # Comprehensive skill for live Word manipulation & triage
│           ├── references/
│           │   └── error-catalog.md        # Complete COM error catalog & agent actions
│           └── SKILL.md                    # Live Word editing procedures and workflows
├── documents/                              # Dedicated working directory for .docx files (gitignored)
│   ├── .backups/                           # Automatic pre-mutation safety backups (latest 20 kept)
│   └── .gitkeep
├── scratch/                                # Scratchpad scripts and temporary data (gitignored)
│   └── .gitkeep                            # Tracked to preserve folder structure
├── src/                                    # Server and engine implementations
│   ├── __init__.py                         # Package declaration
│   ├── server.py                           # FastMCP server entry point (8 tools, backend routing, backups)
│   ├── oxml_engine.py                      # Headless python-docx + low-level lxml engine
│   └── live_engine.py                      # Live COM engine, worker thread, routing, undo/track
├── tests/                                  # Pytest automated test suites
│   ├── __init__.py
│   ├── test_mutations.py                   # Core headless batch composition & DOM mutation tests (3 tests)
│   ├── test_new_tools.py                   # Headless layout, styles, table, hygiene tests (30 tests)
│   ├── test_live_worker.py                 # COM worker thread, timeout, poisoned-state tests (9 tests)
│   ├── test_live_engine.py                 # Live COM integration tests (12 tests)
│   ├── test_gate_2_parity.py               # Live vs Headless feature parity test (1 test)
│   └── test_outline_parity.py              # Exact outline parity and performance tests (4 tests)
├── .gitignore                              # Git exclusion patterns (scratch/*, .agent-info/, documents/*)
├── .python-version                         # Pinned Python version (3.11.9)
├── demo.py                                 # Standalone end-to-end demonstration script
├── pytest.ini                              # Pytest configuration & marker declarations ('live')
├── requirements.txt                        # Runtime and testing dependencies (minimum versions)
├── README.md                               # Public-facing repository documentation
├── WORKSPACE.md                            # Comprehensive workspace architecture guide
└── WORKSPACE_REVIEW.md                     # Current workspace review & status document
```

> *Note: Internal research notes, design plans, and reviews inside `.agent-info/` are excluded locally (`.git/info/exclude` and `.gitignore`) and untracked from Git.*

---

## 3. Dual-Backend Architecture & Routing Logic

Every MCP call routes through `pick_backend(file_path, tool_name, is_mutation)` in [`src/live_engine.py`](src/live_engine.py):

```
Tool Invocation (file_path)
          │
          ▼
   pick_backend()
   ├── 1. Env WORD_BACKEND == "headless" ────────────────────────► Headless
   │      (WORD_BACKEND="live" forces live and raises LiveWordError if not open; see WORKSPACE.md §6)
   ├── 2. pywin32 missing OR Word inactive ──► Check lock file ──► Headless (or reject mutating)
   ├── 3. Document found in Word.Application.Documents ──────────► Live COM
   ├── 4. Document in ProtectedViewWindows ──────────────────────► Raise LiveWordError
   ├── 5. Document is Cloud-Hosted (OneDrive/SharePoint URL) ────► Raise LiveWordError
   ├── 6. Word lock file found (~$*.docx) but unattached ────────► Reject mutating / Headless read
   └── 7. Default fallback ──────────────────────────────────────► Headless
```

### Response Prefix Contract
* All responses from [`src/server.py`](src/server.py) indicate the active backend:
  * String dumps: Prefixed with `[backend: live]\n` or `[backend: headless]\n`.
  * Dictionaries: Annotated with `"backend": "live"` or `"backend": "headless"`.

### COM Concurrency & Safety Safeguards
* **Single Dedicated Worker Thread:** COM operations run inside `ComWorker` with `pythoncom.CoInitialize()`. COM objects never cross thread boundaries.
* **Re-entrancy Protection:** Direct-thread execution bypasses queuing when a helper method already running on the COM worker thread invokes `com_worker.run()`, preventing self-deadlocks.
* **Per-Call Timeout:** Default 20 seconds (configurable via `LIVE_CALL_TIMEOUT`).
* **Poisoned State Protection:** If a call blocks on a Word modal dialog and times out, subsequent calls fail fast immediately rather than queuing behind the stuck thread until the dialog is dismissed.
* **Automatic Pre-Mutation Backups:** Before any mutating operation executes in `src/server.py`, the target document is safely copied to `documents/.backups/` (retaining the 20 most recent revisions).
  * **Backup limitation:** Backups copy the on-disk file. Any unsaved edits made directly in the active Word session prior to the tool call are not captured in the backup until saved to disk by the user.

---

## 4. The 8 MCP Tools Status Matrix & Known Limitations

| # | MCP Tool Name | Headless Engine (`DocxEngine`) | Live Engine (`LiveWordEngine`) | Live Tracking Policy | Primary Function |
|---|---------------|:------------------------------:|:------------------------------:|:--------------------:|------------------|
| 1 | `read_document` | ✅ Full (`outline`, `full`, `table`, `metadata`) | ✅ Full (`outline`, `full`, `table`, `metadata`) | Read-only | 0-based structural dump and outline inspection (bulk range query) |
| 2 | `compose_document` | ✅ Batch writer | ✅ Full COM Batch Writer | Tracked | Sequential creation of headings, paragraphs, lists, tables, and page breaks |
| 3 | `mutate_element` | ✅ Low-level XML DOM | ✅ COM Range DOM | Tracked | Surgical `insert_before`, `insert_after`, `replace`, `delete` |
| 4 | `find_and_replace` | ✅ Scoped search | ✅ Locate-then-set | Tracked | Search and replace in body, tables, headers/footers; skips tracked deletions |
| 5 | `set_page_layout` | ✅ All `<w:sectPr>` sections | ✅ COM PageSetup | Untracked | Sets margins, orientation, paper size (A4, Letter, Legal) across sections |
| 6 | `update_styles` | ✅ ECMA-376 schema order | ✅ COM Styles | Untracked | Updates fonts, sizes, bold, colors, line spacing, space before/after, and `keepNext` |
| 7 | `format_table_grid` | ✅ XML table borders & fills | ✅ COM Table Borders & Fills | Untracked | Standardizes borders, padding, `#E8ECEF` header shading, `cantSplit`, `tblHeader` |
| 8 | `clean_structure` | ✅ Structural hygiene | ✅ COM Structural Hygiene | Mixed | Strips empty paragraphs/tables, caps page breaks, supports `dry_run` |

### Known Limitations & Scope Boundaries
* **`find_and_replace` Scope:** Searches document body, tables, and headers/footers (`scope="all"`, `"body"`, `"tables"`, `"headers_footers"`). Footnotes and endnotes are not searched.
* **`find_and_replace` Dry Run:** `dry_run` is not supported on `find_and_replace` (supported only on `clean_structure`).
* **Replace Formatting:** In live mode, text replacement via `Range.Text` inherits the formatting of the first character of the replaced range.
* **Multiple Word Instances:** Detection attaches to the instance returned by `GetActiveObject`. If Word is running in multiple separate process instances, documents in secondary instances typically trigger the best-effort lock-file safety net rather than direct COM attachment.

---

## 5. Experiment Results & Architectural Decisions

| Experiment | Focus Area | Result / Status | Architectural Decision |
|------------|------------|:---------------:|------------------------|
| **E1** | Revisions & Numbering | Verified | Fully tracked-deleted paragraphs are skipped from `[p:N]` numbering to keep live indices identical to headless. Table paragraphs are addressed exclusively by `[table:N]`. |
| **E2** | Backup Copy | Verified | `shutil.copy2` reliably copies `.docx` files while open in Word on Windows (`FILE_SHARE_READ` granted). |
| **E3** | Lock File Naming | Verified | Word creates owner lock files (`~$<filename>` for short names, or replacing the first two characters with `~$` for longer names, e.g. `~$ve_test.docx`). Used as a best-effort safety net to reject mutating writes when Word is running elevated or in an unattached instance. |
| **E4** | Running Object Table (ROT) | Skipped | Enumerating ROT for multiple instances is complex and brittle across Office versions. The lock-file safety net blocks writes in most cases where Word has the file open but the instance cannot be attached (for example elevated Word or a second instance). It cannot protect a file Word has open without creating a lock file. |
| **E5** | Single Undo Record | Verified | Wrapping COM operations in `doc.Application.UndoRecord` reliably coalesces complex multi-step edits into a single `Ctrl + Z` step. |
| **E6** | Performance Budget (2.0s) | Passed (1.652s) | Outline extraction across 1,000 paragraphs completed in 1.652s (<2.0s budget). Live outline follows the headless document-order rule with body paragraph threshold `OUTLINE_LINE_LIMIT = 40`, leveraging `doc.GetCrossReferenceItems(1)` to index headings in bulk. |
| **E7** | Revision Author | Not Adopted | Setting `Application.UserName` alters global Word user preferences across sessions with risk of leaving modified user state on unexpected process exit. Retained Word's default user author identity. |
| **E8** | `InsertXML` Spike | Not Adopted | COM `Range.InsertXML` produced styling discrepancies and unpredictable undo/track-changes grouping compared to direct COM DOM manipulation. Direct COM manipulation adopted for all tools. |

---

## 6. Test Suite & Verification Metrics

```text
With Microsoft Word running: 59 passed in ~16s
With Microsoft Word closed:  42 passed, 17 skipped
```

| Suite | File | Tests | Status | Scope |
|-------|------|:-----:|:------:|-------|
| Core Headless | [`tests/test_mutations.py`](tests/test_mutations.py) | 3 | **PASSED** | Batch composition, XML paragraph mutations, find-and-replace |
| Advanced Headless | [`tests/test_new_tools.py`](tests/test_new_tools.py) | 30 | **PASSED** | Path normalization, page layout, style order, tables, hygiene |
| COM Worker | [`tests/test_live_worker.py`](tests/test_live_worker.py) | 9 | **PASSED** | Timeouts, poisoned state, thread preservation, lock detection |
| Live Integration | [`tests/test_live_engine.py`](tests/test_live_engine.py) | 12 | **PASSED** | Live open/close, outline, mutations, undo, compose, layout, styles, tables, cleanup |
| Parity Verification | [`tests/test_gate_2_parity.py`](tests/test_gate_2_parity.py) | 1 | **PASSED** | End-to-end parity validation across all 8 tools on `academic-id.yaml` |
| Outline Parity | [`tests/test_outline_parity.py`](tests/test_outline_parity.py) | 4 | **PASSED** | Exact outline parity across short/long docs, tables, deletions, and 1,000-paragraph perf |

---

## 7. Gate Status & Human Sign-off Verification

| Gate | Description | Type | Status | Details / Actions |
|------|-------------|:----:|:------:|-------------------|
| **Gate 0** | Preparation & COM Worker Safety | Automated | **PASSED** | Worker unit tests pass; `HAS_COM=False` headless fallback verified. |
| **Gate 1** | Interactive Live Checklist | **[HUMAN]** | **PASSED** | Verified in active Word session on `documents/live_test.docx`: (1) edit appeared within ~1s, (2) single Ctrl+Z undo verified, (3) tracked changes displayed, (4) headless fallback verified on close, (5) out-of-range clean error, (6) modal dialog safety & recovery verified. |
| **Gate 2** | Tool Parity Comparison | Automated | **PASSED** | `academic-id.yaml` pipeline executed live and headless on identical input; metadata, layout, and styles match. |
| **Gate 3** | Full Test Parity & Dry Run | Automated | **PASSED (scoped: `dry_run` on `clean_structure` only)** | Automated tests passing with Word open (17 skipped when Word closed); `dry_run` verified on `clean_structure` only. |
| **Gate 4** | Final Launch & Merge Review | **[HUMAN]** | **PENDING USER SIGN-OFF** | User review of Draft PR #2, documentation, rules, and skills before merging to `master`. |

### Gate 1 [HUMAN] Checklist (Verified & Signed Off)
Executed on `documents/live_test.docx`:
- [x] 1. Edit appears on screen within ~1 second.
- [x] 2. `Ctrl + Z` undoes exactly one agent tool call.
- [x] 3. The edit appears as a tracked change that can be accepted or rejected in Word.
- [x] 4. After closing Word, the same tool call works headless without errors.
- [x] 5. A deliberately wrong index (e.g., `[p:9999]`) returns a clear out-of-range message without crashing.
- [x] 6. Triggering a tool call while a Word modal dialog (such as *Save As*) is open returns a dialog timeout error within 20s, and the subsequent call succeeds once the dialog is dismissed.

---

## 8. Git Health & PR Status

* **Branch:** `feat/live-word-backend`
* **Draft PR:** #2 (`https://github.com/TheBarli/MsWord-ai-integration/pull/2`)
* **Scratch Directory:** `scratch/` tracked via `scratch/.gitkeep`; contents (`scratch/*`) gitignored.
* **Auxiliary Information:** `.agent-info/` excluded locally in `.git/info/exclude` and `.gitignore`, untracked from Git.
* **Working Tree:** Clean. Commits held locally; awaiting user review before push/merge.
