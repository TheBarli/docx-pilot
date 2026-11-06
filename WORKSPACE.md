# Workspace Architectural Overview: docx-pilot

Welcome to the **docx-pilot** workspace. This repository contains an isolated, production-grade [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server engineered specifically for AI coding assistants (such as Antigravity IDE) to generate, inspect, surgically mutate, and format Microsoft Word (`.docx`) documents with high token efficiency and zero OpenXML schema corruption.

---

## 1. High-Level Mission & Design Principles

Standard automated `.docx` workflows (like plain `python-docx`) suffer from three major issues when used in LLM-assisted environments:
1. **Append-Only Limitations**: Standard high-level libraries lack easy primitives for inserting or deleting elements in the middle of a document without rewriting the entire body.
2. **Token Bloat & Hallucination**: Exposing dozens of granular tools burns context window tokens and causes off-by-one errors in document navigation.
3. **XML Schema Corruption**: Programmatic XML injection frequently violates the strict ECMA-376 OpenXML schema ordering, causing Microsoft Word to display "Word found unreadable content / Word experienced an error trying to open the file" repair dialogs.

This workspace solves those problems using three foundational pillars:
- **Consolidated 8-Tool API**: Combines document inspection, batch creation, surgical DOM mutation, global search-and-replace, section layout, style configuration, table standardization, and structural hygiene into 8 high-density MCP endpoints.
- **Strict 0-Based Indexing**: All paragraphs (`[p:0]`, `[p:1]`) and tables (`[table:0]`) use explicit 0-based coordinate tags in inspection outputs and mutation inputs, preventing off-by-one errors.
- **Low-Level OpenXML Engine with ECMA-376 Ordering**: Directly manipulates the underlying `lxml` nodes (`addprevious`, `addnext`, `parent.remove`) and enforces strict schema child sequence ordering (`STYLE_ORDER`, `PPR_ORDER`, `RPR_ORDER`) so documents open cleanly in Microsoft Word.

---

## 2. Workspace Directory Structure

```text
docx-pilot/
├── .agents/                                # Antigravity agent configuration & presets
│   ├── mcp_config.json                     # Portable workspace-level MCP registration
│   ├── presets/                            # Standardized document formatting profiles (YAML)
│   │   ├── academic-id.yaml                # Standard Indonesian academic format (A4, 4-4-3-3 cm)
│   │   ├── academic-id-advanced.yaml       # Advanced academic preset with H1–H5 & pagination
│   │   └── clean-uniform.yaml              # General business clean styling profile
│   ├── rules/                              # Agent behavior and orchestration rules
│   │   ├── document-styling.md             # Rule orchestrating the 5-step styling pipeline
│   │   ├── live-word-mode.md               # Rule governing live Word backend usage & constraints
│   │   └── [local dev rules excluded]      # Git, GH, and debugging rules via .git/info/exclude
│   └── skills/
│       └── live-word-editing/              # Skill containing live editing workflow & error triage
│           ├── references/
│           │   └── error-catalog.md        # Comprehensive COM error catalog & recovery actions
│           └── SKILL.md                    # Live Word editing playbook and procedures
├── documents/                              # Dedicated working folder for input/output .docx files
│   ├── .backups/                           # Automatic pre-mutation safety backups (latest 20 retained)
│   └── .gitkeep                            # Keeps directory tracked while ignoring binary docs
├── scratch/                                # Scratchpad scripts and temporary data (gitignored)
│   └── .gitkeep                            # Keeps directory tracked in repository
├── src/                                    # Server application source code
│   ├── __init__.py                         # Package declaration
│   ├── core/                               # Core shared infrastructure
│   │   ├── __init__.py
│   │   ├── constants.py                    # Schema ordering, namespaces, limits, PROJECT_ROOT
│   │   ├── exceptions.py                   # LiveWordError & error message catalog
│   │   └── paths.py                        # Multi-candidate path normalization
│   ├── headless/                           # OpenXML headless engine (python-docx + lxml)
│   │   ├── __init__.py
│   │   ├── xml_helpers.py                  # ECMA-376 schema sort & document creation
│   │   ├── tools_read.py                   # read_document headless mixin
│   │   ├── tools_edit.py                   # compose, mutate, find_and_replace, debloat mixin
│   │   ├── tools_format.py                 # page layout, styles, tables, cleanup mixin
│   │   └── engine.py                       # DocxEngine assembling mixins
│   ├── live/                               # Live Microsoft Word COM backend
│   │   ├── __init__.py
│   │   ├── worker.py                       # STA ComWorker thread & com_worker singleton
│   │   ├── undo.py                         # UndoRecord & live_call context managers
│   │   ├── router.py                       # Backend detection & pick_backend routing
│   │   ├── indexer.py                      # 0-based coordinate indexing & outline helpers
│   │   ├── tools_read.py                   # read_document COM mixin
│   │   ├── tools_edit.py                   # compose, mutate, find_and_replace COM mixin
│   │   ├── tools_format.py                 # page layout, styles, tables, cleanup COM mixin
│   │   └── engine.py                       # LiveWordEngine assembling mixins
│   ├── oxml_engine.py                      # Backward-compatibility facade for headless engine
│   ├── live_engine.py                      # Backward-compatibility facade for live engine
│   └── server.py                           # FastMCP server entry point exporting the 8 tools
├── tests/                                  # Pytest automated verification suite
│   ├── __init__.py
│   ├── test_mutations.py                   # Tests for compose, read, and low-level DOM mutations
│   ├── test_new_tools.py                   # Tests for page layout, styles, tables, and hygiene
│   ├── test_live_worker.py                 # Single-threaded COM worker, timeouts, and state safety
│   ├── test_live_engine.py                 # Live Word COM integration tests (12 tests)
│   ├── test_gate_2_parity.py               # End-to-end parity test across both backends
│   └── test_outline_parity.py              # Tests for live/headless outline parity and performance
├── .gitignore                              # Git exclusions (binaries, caches, scratch files)
├── .python-version                         # Pinned Python version (3.11.9)
├── demo.py                                 # Self-contained end-to-end demo script
├── pytest.ini                              # Pytest configuration and live test marker registration
├── requirements.txt                        # Runtime and test dependencies (minimum versions)
├── README.md                               # Public-facing repository documentation
├── WORKSPACE.md                            # Comprehensive workspace architecture guide
└── WORKSPACE_REVIEW.md                     # Workspace audit and gate verification status
```

---

## 3. Core Engine Architecture (`src/`)

The engine source is organized into modular packages targeting $\le 500$ lines per file, with backward-compatibility facades at the root of `src/`.

### Shared Infrastructure (`src/core/`)
- **`constants.py`**: ECMA-376 schema child ordering sequences (`STYLE_ORDER`, `PPR_ORDER`, `RPR_ORDER`), XML namespaces (`W_NS`), outline truncation limit (`OUTLINE_LINE_LIMIT`), and dynamic `PROJECT_ROOT` detection.
- **`exceptions.py`**: `LiveWordError` hierarchy and centralized `ERROR_MESSAGES` catalog.
- **`paths.py`**: Multi-candidate path normalizer (`normalize_path`, `_norm`, `paths_match`) supporting absolute paths, working-directory relative paths, and project-root relative paths.

### Headless OpenXML Backend (`src/headless/`)
- **`xml_helpers.py`**: ECMA-376 schema sorting (`reorder_children`) and document initialization (`get_or_create_document`).
- **`tools_read.py` (`ReadTools`)**: Headless `read_document` implementing outline, full, table, and metadata modes.
- **`tools_edit.py` (`EditTools`)**: Surgical DOM mutations (`mutate_element`, `find_and_replace`, `compose_document`, `debloat_document`).
- **`tools_format.py` (`FormatTools`)**: OpenXML styling (`set_page_layout`, `update_styles`, `format_table_grid`, `clean_structure`).
- **`engine.py` (`DocxEngine`)**: Assembles `DocxEngine(ReadTools, EditTools, FormatTools)` with zero behavior change.

### Live Microsoft Word COM Backend (`src/live/`)
- **`worker.py`**: Single-threaded STA `ComWorker` thread, com error detection, and `com_worker` singleton.
- **`undo.py`**: Atomic Word `Application.UndoRecord` context manager (`live_call`) and Track Changes scoping (`tracking`).
- **`router.py`**: Dynamic backend routing (`pick_backend`, `find_open_document`, `has_word_lock_file`).
- **`indexer.py`**: Document-order 0-based coordinate mapping (`_build_visible_map`, `_visible_map`, `_com_idx`, `_get_local_outline`).
- **`tools_read.py` (`ReadTools`)**: Live COM `read_document` with fast heading detection.
- **`tools_edit.py` (`EditTools`)**: Live COM mutations (`mutate_element`, `find_and_replace`, `compose_document`) wrapped in UndoRecord.
- **`tools_format.py` (`FormatTools`, `_hex_to_bgr`)**: Live COM styling (`set_page_layout`, `update_styles`, `format_table_grid`, `clean_structure`).
- **`engine.py` (`LiveWordEngine`)**: Assembles `LiveWordEngine(ReadTools, EditTools, FormatTools)` with zero behavior change.

### Backward-Compatibility Facades
- **`src/oxml_engine.py`**: Re-exports all symbols from `src.core` and `src.headless` with explicit `__all__`.
- **`src/live_engine.py`**: Re-exports all symbols from `src.core` and `src.live` with explicit `__all__`.

### FastMCP Server (`src/server.py`)
The FastMCP entry point running over standard I/O (`stdio`). It sets the working directory to `PROJECT_ROOT`, manages automatic document backups, dynamically routes calls between headless and live backends via `pick_backend()`, and exposes the 8 MCP tools.

---

## 4. The 8 MCP Tools

| # | Tool Name | Primary Purpose | Key Parameters |
|---|-----------|-----------------|----------------|
| 1 | [`read_document`](src/server.py) | Inspect document structure, outline, tables, or metadata | `file_path`, `mode` (`"outline"`, `"full"`, `"table"`, `"metadata"`), `table_index` |
| 2 | [`compose_document`](src/server.py) | Batch creation of headings, paragraphs, lists, tables, and page breaks | `file_path`, `actions` (list of element dicts), `create_if_missing` |
| 3 | [`mutate_element`](src/server.py) | Surgical DOM insertions, deletions, or in-place replacements at `[p:N]` | `file_path`, `target_index`, `action` (`"insert_before"`, `"insert_after"`, `"delete"`, `"replace"`), `content` |
| 4 | [`find_and_replace`](src/server.py) | Text search-and-replace across paragraphs and tables (preserves run formatting on headless; inherits first character formatting on live) | `file_path`, `find_text`, `replace_text`, `match_case`, `scope` |
| 5 | [`set_page_layout`](src/server.py) | Applies margins and paper size to **all** `<w:sectPr>` elements in the XML | `file_path`, `top_cm`, `bottom_cm`, `left_cm`, `right_cm`, `page_size`, `orientation` |
| 6 | [`update_styles`](src/server.py) | Configures font family, font size, line spacing, spacing, and keepNext in `word/styles.xml` | `file_path`, `styles_config` (mapping of style names to formatting attributes) |
| 7 | [`format_table_grid`](src/server.py) | Normalizes table borders, header shading, cell padding, `cantSplit`, and `tblHeader` | `file_path`, `table_indices`, `border_sz`, `border_color`, `header_shading`, `cant_split`, `repeat_header`, `font_name`, `font_size_pt` |
| 8 | [`clean_structure`](src/server.py) | Removes empty paragraphs, empty tables, extracts single-image tables, and caps breaks | `file_path`, `remove_empty_paragraphs`, `remove_empty_tables`, `extract_single_image_tables`, `max_consecutive_breaks`, `normalize_all_colors_black`, `dry_run` |

---

## 5. Document Presets & Styling Pipeline (`.agents/presets/`)

Formatting configurations are defined in portable YAML presets that map directly to the parameters of the MCP styling tools:

1. **[`academic-id.yaml`](.agents/presets/academic-id.yaml)**:
   - Paper Size: **A4** (210 × 297 mm) portrait.
   - Margins: **4.0 cm top, 3.0 cm bottom, 4.0 cm left, 3.0 cm right** (Indonesian standard).
   - Typography: Times New Roman 12pt, 1.5 line spacing, justified, 0pt space after.
   - Headings: Times New Roman 12pt bold, `keep_next: true`.
   - Tables: 0.5pt black border, `F2F2F2` header fill, row anti-split enabled.
   - Hygiene: Full structural cleanup enabled, black font color normalization.

2. **[`academic-id-advanced.yaml`](.agents/presets/academic-id-advanced.yaml)**:
   - Extends the base academic preset with fine-grained hierarchy (H1 through H5), bullet list indent control, and pagination management.

3. **[`clean-uniform.yaml`](.agents/presets/clean-uniform.yaml)**:
   - General corporate/modern document profile with uniform 2.54 cm (1 inch) margins and Calibri typography.

### Orchestration Sequence
When formatting documents using presets, the agent follows the pipeline defined in [`.agents/rules/document-styling.md`](.agents/rules/document-styling.md):
```text
1. set_page_layout  -->  Apply page size, orientation, and margins across all sections
2. update_styles    -->  Update word/styles.xml with fonts, line spacing, and keepNext
3. format_table_grid-->  Enforce uniform borders, header shading, and anti-split rows
4. clean_structure  -->  Purge trailing empty paragraphs, empty tables, and rogue page breaks
5. read_document    -->  Verify output metadata (paragraphs, tables, styles)
```

---

## 6. Dual Backend Engine Architecture (Live & Headless)

The server supports transparent dual-backend operation:

1. **Headless Engine ([`src/oxml_engine.py`](src/oxml_engine.py))**:
   - Operates on static files on disk via `python-docx` and `lxml`.
   - Low-level XML tree mutations with strict ECMA-376 schema ordering.
   - Non-destructive `dry_run: bool = False` for structural hygiene.

2. **Live Microsoft Word Engine ([`src/live_engine.py`](src/live_engine.py))**:
   - Automatically detects if the document is active in Microsoft Word using `GetActiveObject("Word.Application")`.
   - Interacts live via Windows COM with zero external window spawning; the live engine never launches Word, only attaching to an already running instance.
   - Groups multi-step actions inside Word's `Application.UndoRecord` so that each live tool call is a single undo step (`Ctrl + Z`); content edits are tracked, pure formatting edits are untracked.
   - Visual Track Changes for mutations and text replacements.
   - Dedicated single-threaded worker (`ComWorker`) with 20s timeout and poisoned-state protection.
   - Re-entrancy protection and Windows desktop station bridging (`SetThreadDesktop("Default")`).

3. **Automatic Safety Backups**:
   - Before any mutating tool execution, a timestamped copy is saved to `documents/.backups/`.
   - Automatically rotates and retains the latest `BACKUP_KEEP = 20` copies per file stem.
   - **Limitation note:** Backups copy the on-disk file. Any unsaved edits made in an open Word session prior to the agent call are not captured in the backup until saved to disk by the user.

4. **Scope & Tool Limitations**:
   - `find_and_replace`: Searches body text, tables, and headers/footers (`scope="all"` or `scope="headers_footers"`). Footnotes and endnotes are not searched.
   - `find_and_replace` Formatting: Replaced text preserves granular run formatting on the headless backend; on the live backend, replaced text inherits the character formatting of the first replaced character.
   - `dry_run`: Supported on `clean_structure`; not supported on `find_and_replace`.

---

## 7. Antigravity IDE Integration & Local Setup

### Workspace Configuration
The workspace is pre-configured with [`.agents/mcp_config.json`](.agents/mcp_config.json):
```json
{
  "mcpServers": {
    "word-docx": {
      "command": "python",
      "args": ["src/server.py"],
      "env": {
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```
When opened in Antigravity IDE, the `word-docx` server is automatically recognized and loaded into the agent's toolset.

### Rules & Skills
- **Rule [`live-word-mode`](.agents/rules/live-word-mode.md)**: Governs live vs headless tool invariants, outline reuse, and safety constraints.
- **Rule [`document-styling`](.agents/rules/document-styling.md)**: Enforces the 5-step preset styling pipeline.
- **Skill [`live-word-editing`](.agents/skills/live-word-editing/SKILL.md)**: Procedure and error catalog lookup for interactive live editing.

### Environment Variables

| Variable | Values / Default | Effect |
|----------|------------------|--------|
| `WORD_BACKEND` | `auto` (default), `live`, `headless` | Forces or auto-detects the backend. `live` with the file not open in Word returns an error. |
| `LIVE_CALL_TIMEOUT` | seconds, default `20` | Per-call COM timeout before the worker is marked poisoned. |

### Python Environment
- Python version is pinned to **3.11.9** via [`.python-version`](.python-version).
- Production dependencies in [`requirements.txt`](requirements.txt) (minimum versions):
  - `fastmcp>=4.0.0`
  - `python-docx>=1.2.0`
  - `pydantic>=2.0.0`
  - `lxml>=6.1.0`
  - `pywin32>=306; sys_platform == 'win32'`
- Testing dependencies:
  - `pytest>=9.0.0`

---

## 8. Testing & Quality Assurance

Test execution behavior:
- **With Word running:** 59 passed (42 headless, 17 live tests).
- **With Word closed:** 42 passed, 17 live tests skipped.

```powershell
# Run headless tests (no Word required)
.venv\Scripts\python.exe -m pytest -m "not live" -v

# Run full test suite including live COM tests (with Word open)
.venv\Scripts\python.exe -m pytest tests/ -v
```

### Test Suites:
- **[`tests/test_mutations.py`](tests/test_mutations.py)**: Headless batch creation, reading, and XML paragraph mutations (3 tests).
- **[`tests/test_new_tools.py`](tests/test_new_tools.py)**: Headless layout, styles, table formatting, and structural hygiene with dry_run (30 tests).
- **[`tests/test_live_worker.py`](tests/test_live_worker.py)**: Single-threaded COM worker isolation, timeouts, poisoned state, path matching, and lock file detection (9 tests).
- **[`tests/test_live_engine.py`](tests/test_live_engine.py)**: 12 live COM tests verifying all 8 tools, single-step UndoRecord, and Track Changes.
- **[`tests/test_gate_2_parity.py`](tests/test_gate_2_parity.py)**: Full end-to-end parity test executing `academic-id.yaml` styling pipeline across both backends on identical input (1 test).
- **[`tests/test_outline_parity.py`](tests/test_outline_parity.py)**: 4 live COM tests verifying outline parity across short docs, 40+ line docs, tables, and tracked deletions under 2.0s budget (4 tests).

---

## 9. Practical Demo

To see the server and engine in action end-to-end, execute [`demo.py`](demo.py):
```powershell
.venv\Scripts\python.exe demo.py
```
This script creates a rich formatted document in `documents/demo_document.docx`, inspects the 0-based outline, inserts content via low-level XML manipulation, executes targeted find-and-replace, and verifies the final structure.
