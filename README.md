# docx-pilot: Word DOCX FastMCP Server for Antigravity IDE

A high-performance, token-efficient Model Context Protocol (MCP) server for inspecting, generating, and surgically mutating Microsoft Word (`.docx`) documents directly from Antigravity IDE.

Featuring a **transparent dual-backend architecture**, `docx-pilot` automatically detects whether a document is actively open in Microsoft Word on Windows to perform real-time, visual live editing with Track Changes and single-step `Ctrl + Z` undo, while providing identical schema-safe manipulation when running headless.

---

## Highlights

* **Consolidated 8-Tool Architecture**: Covers the full document manipulation lifecycle—outline/full inspection, batch generation, surgical XML DOM mutation, global/scoped find-and-replace, section-level page layout, ECMA-376 style configuration, table grid normalization, and automated structural hygiene.
* **Transparent Dual-Backend (Live COM + Headless)**:
  * **Live Backend (`LiveWordEngine`)**: Attaches via Windows COM (`pywin32`) to the active Microsoft Word desktop session. Live edits appear on screen in ~1 second. Content edits are logged under Word's native Track Changes, and multi-step tool calls are coalesced into single-step `Ctrl + Z` undo records (`UndoRecord`).
  * **Headless Backend (`DocxEngine`)**: Operates statically on disk via `python-docx` and `lxml` with strict ECMA-376 XML child element ordering. Operates without requiring Microsoft Word installed or running.
* **COM Concurrency & Safety Nets**: Dedicated single-threaded worker (`ComWorker`) with interactive desktop thread binding, 20s timeout against modal dialogs, poisoned-state fast-failing, and best-effort lock-file detection (`~$*`).
* **Automatic Safety Backups**: Creates timestamped pre-mutation backups in `documents/.backups/` before any write operation, retaining the latest 20 revisions per file stem.
* **Strict 0-Based Coordinate Indexing**: All paragraph (`[p:N]`), table (`[table:N]`), row, and column identifiers strictly use 0-based indexing to eliminate LLM off-by-one hallucinations.
* **Document-Order Outline Parity**: High-speed outline extraction interleaving headings, body paragraphs, and tables in document order with a shared `OUTLINE_LINE_LIMIT = 40` cap and sub-2.0s performance across 1,000 paragraphs.
* **ECMA-376 Schema Compliance**: Automated child sequence sorting (`STYLE_ORDER`, `PPR_ORDER`, `RPR_ORDER`) ensures manipulated documents open cleanly in Microsoft Word without "repaired content" warnings.
* **Pinned Python 3.11 Runtime**: Guarantees pre-built C-wheels on Windows for `lxml`, `python-docx`, and `pywin32`.

---

## Directory Structure

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
│   │   └── live-word-mode.md               # Rule governing live Word backend usage & constraints
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
│   ├── oxml_engine.py                      # Core python-docx + lxml DOM manipulation engine
│   ├── live_engine.py                      # Live Microsoft Word COM backend & worker thread
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

## Dual-Backend Architecture & Routing

Every MCP tool call routes automatically through `pick_backend(file_path, tool_name, is_mutation)` in `src/live_engine.py`:

```text
Tool Invocation (file_path)
          │
          ▼
   pick_backend()
   ├── 1. Env WORD_BACKEND == "headless" ────────────────────────► Headless
   │      (WORD_BACKEND="live" forces live; raises if not open)
   ├── 2. pywin32 missing OR Word inactive ──► Check lock file ──► Headless (or reject mutating)
   ├── 3. Document found in Word.Application.Documents ──────────► Live COM
   ├── 4. Document in ProtectedViewWindows ──────────────────────► Raise LiveWordError
   ├── 5. Document is Cloud-Hosted (OneDrive/SharePoint URL) ────► Raise LiveWordError
   ├── 6. Word lock file found (~$*.docx) but unattached ────────► Reject mutating / Headless read
   └── 7. Default fallback ──────────────────────────────────────► Headless
```

### Response Contract
All responses identify the backend that processed the operation:
* **String dumps**: Prefixed with `[backend: live]\n` or `[backend: headless]\n`.
* **JSON / Dictionaries**: Annotated with `"backend": "live"` or `"backend": "headless"`.

### Live Tracking & Undo Policies
| MCP Tool | Live Tracking Mode | Undo Scope | Notes |
|:---|:---:|:---:|:---|
| `read_document` | Read-only | N/A | Excludes tracked deletions from numbering |
| `compose_document` | Untracked | 1 `Ctrl + Z` step | Batch creation wrapped in `UndoRecord` |
| `mutate_element` | **Tracked** | 1 `Ctrl + Z` step | In-place insertions/deletions logged in revision history |
| `find_and_replace` | **Tracked** | 1 `Ctrl + Z` step | Text replacement inherits leading run style |
| `set_page_layout` | Untracked | 1 `Ctrl + Z` step | Section margins & dimensions updated in bulk |
| `update_styles` | Untracked | 1 `Ctrl + Z` step | Document-wide style sheet updates |
| `format_table_grid` | Untracked | 1 `Ctrl + Z` step | Anti-split, repeat headers, borders, shading |
| `clean_structure` | Mixed | 1 `Ctrl + Z` step | Empty elements removed; supports `dry_run` |

---

## Available MCP Tools

### 1. `read_document`
Inspects document structure, outline, tables, or metadata.
* **Signature**:
  ```python
  read_document(
      file_path: str,
      mode: Literal["outline", "full", "table", "metadata"] = "full",
      table_index: Optional[int] = None  # 0-based
  ) -> str | dict
  ```
* **Modes**:
  * `"outline"`: Document-order summary of headings, body paragraphs (up to 40 lines), and tables with explicit tags (`[p:0]`, `[table:0]`). Excludes tracked deletions.
  * `"full"`: Complete line-by-line document dump with `[p:N]` and `[table:N]` coordinate annotations.
  * `"table"`: Extracts table at 0-based `table_index` as formatted Markdown and structured 2D matrix.
  * `"metadata"`: Core properties (author, title, revision, word count, timestamps).

### 2. `compose_document`
Appends a batch of structured elements in a single round-trip.
* **Signature**:
  ```python
  compose_document(
      file_path: str,
      actions: list[dict],
      create_if_missing: bool = True
  ) -> dict
  ```
* **Supported Action Types**:
  * `{"type": "heading", "text": "...", "level": 1}` (levels 1–9)
  * `{"type": "paragraph", "text": "...", "style": "Normal", "bold": False, "italic": False}`
  * `{"type": "list_item", "text": "...", "style": "List Bullet"}`
  * `{"type": "table", "headers": ["Col 0", "Col 1"], "rows": [["val0", "val1"]], "style": "Table Grid"}`
  * `{"type": "page_break"}`

### 3. `mutate_element`
Performs surgical XML/COM mutations at a 0-based paragraph index (`[p:N]`).
* **Signature**:
  ```python
  mutate_element(
      file_path: str,
      target_index: int,  # 0-based paragraph index matching [p:N]
      action: Literal["insert_before", "insert_after", "delete", "replace"],
      content: Optional[dict] = None
  ) -> dict
  ```
* **Actions**:
  * `"insert_before"`: Inserts element immediately preceding `[p:target_index]`.
  * `"insert_after"`: Inserts element immediately following `[p:target_index]`.
  * `"delete"`: Deletes the target paragraph cleanly without leaving empty text artifacts.
  * `"replace"`: Replaces paragraph text/style in place.

### 4. `find_and_replace`
Global or scoped text replacement across the document body, tables, and headers/footers.
* **Signature**:
  ```python
  find_and_replace(
      file_path: str,
      find_text: str,
      replace_text: str,
      match_case: bool = True,
      scope: Literal["all", "body", "tables", "headers_footers"] = "all"
  ) -> dict
  ```

### 5. `set_page_layout`
Modifies margins and page dimensions across **all** `<w:sectPr>` sections in the document.
* **Signature**:
  ```python
  set_page_layout(
      file_path: str,
      top_cm: float = 2.54,
      bottom_cm: float = 2.54,
      left_cm: float = 2.54,
      right_cm: float = 2.54,
      page_size: Literal["letter", "a4", "legal"] = "letter",
      orientation: Literal["portrait", "landscape"] = "portrait"
  ) -> dict
  ```

### 6. `update_styles`
Directly edits document styles (`word/styles.xml` headless or `Document.Styles` live) to configure font families, font sizes, line spacing, space before/after, colors, and `keepNext` pagination rules.
* **Signature**:
  ```python
  update_styles(
      file_path: str,
      styles_config: dict[str, dict]
  ) -> dict
  ```

### 7. `format_table_grid`
Standardizes borders, row pagination properties (`cantSplit`), header repeats (`tblHeader`), header shading, and padding across tables.
* **Signature**:
  ```python
  format_table_grid(
      file_path: str,
      table_indices: Optional[list[int]] = None,  # None targets all tables
      border_sz: int = 4,                         # 1/8 pt units (4 = 0.5 pt)
      border_color: str = "000000",
      header_shading: Optional[str] = "F2F2F2",
      cant_split: bool = True,                    # Prevents rows splitting across pages
      repeat_header: bool = True,                 # Repeats header row on multi-page tables
      font_name: Optional[str] = "Times New Roman",
      font_size_pt: Optional[float] = 11.0
  ) -> dict
  ```

### 8. `clean_structure`
Performs automated hygiene passes to eliminate structural bloat:
* Strips trailing empty paragraphs that contain no text, drawings, or breaks.
* Deletes empty tables containing no text or images.
* Extracts single images wrapped in tables and converts them to centered standalone paragraphs.
* Caps consecutive page breaks.
* Optionally normalizes all body text colors to pure black (`#000000`).
* Supports non-destructive preview inspection via `dry_run=True`.
* **Signature**:
  ```python
  clean_structure(
      file_path: str,
      remove_empty_paragraphs: bool = True,
      remove_empty_tables: bool = True,
      extract_single_image_tables: bool = True,
      max_consecutive_breaks: int = 1,
      normalize_all_colors_black: bool = False,
      dry_run: bool = False
  ) -> dict
  ```

---

## Formatting Presets & Agent Automation

The repository includes pre-built document styling presets under `.agents/presets/`:
* **`academic-id.yaml`**: Standard Indonesian academic style (A4, 4-4-3-3 cm margins, Times New Roman 12pt, 1.5 line spacing, 0.5pt table borders).
* **`academic-id-advanced.yaml`**: Extended academic profile with fine-grained heading hierarchies (H1–H5), list indentation, and strict pagination.
* **`clean-uniform.yaml`**: General-purpose business profile (Letter/A4, uniform 2.54 cm margins, Calibri typography).

### Sequential Styling Pipeline
When formatting a document according to a preset, execute tools in this sequence:
1. `set_page_layout` $\rightarrow$ Apply section margins, paper size, and orientation.
2. `update_styles` $\rightarrow$ Apply font families, sizes, line spacing, and paragraph rules.
3. `format_table_grid` $\rightarrow$ Enforce borders, header shading, and anti-split rows.
4. `clean_structure` $\rightarrow$ Run with `dry_run=True` for preview, then `dry_run=False` to purge empty nodes.
5. `read_document(mode="metadata")` $\rightarrow$ Verify output metadata and structural counts.

---

## Environment Variables

| Variable | Values / Default | Effect |
|:---|:---|:---|
| `WORD_BACKEND` | `auto` (default), `live`, `headless` | Controls backend routing. When set to `live`, raises `LiveWordError` if the document is not currently open in Word. |
| `LIVE_CALL_TIMEOUT` | seconds, default `20` | Per-call COM timeout before marking the worker poisoned to protect against frozen modal dialogs. |

---

## Antigravity IDE Integration

### Option A: Workspace MCP Registration (Automatic)
The project includes `.agents/mcp_config.json`. When opening `docx-pilot` in Antigravity IDE, the server is automatically discovered and loaded:

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

### Option B: Global Configuration
To make this server accessible across all workspaces in Antigravity IDE, add an entry to `~/.gemini/config/mcp_config.json`:

```json
{
  "mcpServers": {
    "word-docx": {
      "command": "<repo-path>/.venv/Scripts/python.exe",
      "args": ["<repo-path>/src/server.py"],
      "env": {
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

---

## Testing & Quality Assurance

The test suite provides 100% automated verification for both headless operations and live COM integration:

```powershell
# Run full suite (with Word active: 59 passed; with Word closed: 42 passed, 17 skipped)
.venv\Scripts\python.exe -m pytest tests/ -v

# Run headless-only tests
.venv\Scripts\python.exe -m pytest tests/ -m "not live" -v

# Run end-to-end self-contained demo
.venv\Scripts\python.exe demo.py
```

### Verification Gate Summary
* **Gate 0:** COM worker safety, timeout, and headless fallback verified.
* **Gate 1:** Interactive live Word checklist signed off by the repo owner (2026-11-06).
* **Gate 2:** Full styling pipeline parity between headless and live engines verified.
* **Gate 3:** 59 automated test cases passing with scoped `dry_run` verification.
* **Gate 4:** Final launch review and merge into `main` (2026-11-06).
