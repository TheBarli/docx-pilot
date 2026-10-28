# Word DOCX FastMCP Server for Antigravity IDE

A high-performance, token-efficient Model Context Protocol (MCP) server for inspecting, generating, and surgically mutating Microsoft Word (`.docx`) documents directly from Antigravity IDE.

---

## Highlights

* **Consolidated 8-Tool Architecture**: Covers the full document manipulation lifecycle—batch generation, surgical DOM mutation, global find-and-replace, section-level page layout, ECMA-376 style configuration, table grid normalization, and automated structural hygiene.
* **Strict 0-Based Indexing**: All paragraph (`[p:N]`), table (`[table:N]`), row, and column identifiers strictly use 0-based indexing to eliminate off-by-one hallucinations.
* **Low-Level XML Mutations**: Leverages underlying `lxml`/`oxml` element nodes (`addprevious`, `addnext`, `parent.remove`) for in-place paragraph/table insertions and clean deletions without empty-node document corruption.
* **ECMA-376 Schema Compliance**: Automated child sequence sorting (`STYLE_ORDER`, `PPR_ORDER`, `RPR_ORDER`) ensures manipulated documents open cleanly in Microsoft Word without "repaired content" warnings.
* **Pinned Python 3.11 Runtime**: Guarantees pre-built C-wheels on Windows for `lxml` and `python-docx`.

---

## File Structure

```text
MsWord-ai-integration/
├── .agents/
│   ├── presets/                        # Formatting profiles (academic-id, clean-uniform)
│   ├── rules/                          # Document styling orchestration rules
│   └── mcp_config.json                 # Portable workspace MCP config
├── src/
│   ├── __init__.py
│   ├── server.py                       # FastMCP server entry point (8 tools)
│   └── oxml_engine.py                  # Core python-docx + OpenXML mutation engine
├── tests/
│   ├── __init__.py
│   ├── test_mutations.py               # Unit tests for core compose, read, and DOM mutations
│   └── test_new_tools.py               # Unit tests for layout, styles, tables, and hygiene
├── documents/                          # Working directory for .docx files (gitignored)
├── demo.py                             # End-to-end demonstration script
├── .python-version                     # Pinned to 3.11
├── requirements.txt                    # Pinned dependencies
└── README.md
```

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
  * `"outline"`: Overview of headings and tables with explicit 0-based tags (`[p:0]`, `[p:1]`, `[table:0]`).
  * `"full"`: Complete line-by-line document dump with `[p:N]` and `[table:N]` annotations.
  * `"table"`: Extracts table at 0-based `table_index` as formatted Markdown and structured matrix.
  * `"metadata"`: OpenXML properties (author, title, revision, timestamps).

### 2. `compose_document`
Appends a batch of formatted elements in a single round-trip.
* **Signature**:
  ```python
  compose_document(
      file_path: str,
      actions: list[dict],
      create_if_missing: bool = True
  ) -> dict
  ```
* **Action Types**:
  * `{"type": "heading", "text": "...", "level": 1}` (levels 1–9)
  * `{"type": "paragraph", "text": "...", "style": "Normal", "bold": False, "italic": False}`
  * `{"type": "list_item", "text": "...", "style": "List Bullet"}`
  * `{"type": "table", "headers": ["Col 0", "Col 1"], "rows": [["val0", "val1"]], "style": "Table Grid"}`
  * `{"type": "page_break"}`

### 3. `mutate_element`
Performs surgical XML mutations at a 0-based paragraph index (`[p:N]`).
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
  * `"insert_before"`: Inserts element before `[p:target_index]` via `target_p._p.addprevious(...)`.
  * `"insert_after"`: Inserts element after `[p:target_index]` via `target_p._p.addnext(...)`.
  * `"delete"`: Deletes paragraph from DOM via `target_p._p.getparent().remove(target_p._p)`.
  * `"replace"`: Replaces paragraph text/style in place.

### 4. `find_and_replace`
Global or scoped text replacement across the document with run-level formatting preservation.
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
Modifies margins and page dimensions across **all** `<w:sectPr>` elements in `document.xml` (ensuring multi-section documents are uniformly sized).
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
Directly edits `word/styles.xml` to configure font families, font sizes, line spacing, space before/after, colors, and `keepNext` rules. Children are reordered according to the ECMA-376 sequence specification.
* **Signature**:
  ```python
  update_styles(
      file_path: str,
      styles_config: dict[str, dict]
  ) -> dict
  ```

### 7. `format_table_grid`
Standardizes borders, row pagination properties, header fills, and cell margins across tables.
* **Signature**:
  ```python
  format_table_grid(
      file_path: str,
      table_indices: Optional[list[int]] = None,  # None targets all tables
      border_sz: int = 4,                         # 1/8 pt units (4 = 0.5 pt)
      border_color: str = "000000",
      header_shading: Optional[str] = "F2F2F2",
      cant_split: bool = True,                    # Prevents split across page breaks
      repeat_header: bool = True,                 # Repeats header on multi-page tables
      font_name: Optional[str] = "Times New Roman",
      font_size_pt: Optional[float] = 11.0
  ) -> dict
  ```

### 8. `clean_structure`
Performs automated hygiene passes to eliminate structural bloat:
* Strips trailing empty paragraphs (`<w:p>`) that contain no text, drawings, or breaks.
* Deletes empty tables with no text or images.
* Extracts single images wrapped in tables and converts them to centered standalone paragraphs.
* Caps consecutive page breaks.
* Optionally forces all text color declarations in `document.xml` to pure black (`#000000`).
* **Signature**:
  ```python
  clean_structure(
      file_path: str,
      remove_empty_paragraphs: bool = True,
      remove_empty_tables: bool = True,
      extract_single_image_tables: bool = True,
      max_consecutive_breaks: int = 1,
      normalize_all_colors_black: bool = False
  ) -> dict
  ```

---

## Formatting Presets & Agent Automation

The repository includes pre-built document styling presets under `.agents/presets/`:
* **`academic-id.yaml`**: Standard Indonesian academic style (A4, 4-4-3-3 cm margins, Times New Roman 12pt, 1.5 line spacing, 0.5pt table borders).
* **`academic-id-advanced.yaml`**: Extended academic profile with fine-grained heading hierarchies (H1–H5), list indentation, and strict pagination.
* **`clean-uniform.yaml`**: General-purpose business profile (Letter/A4, uniform 2.54 cm margins, Calibri typography).

### Sequential Pipeline Order
When applying a preset, execute tools in this sequence:
1. `set_page_layout`
2. `update_styles`
3. `format_table_grid`
4. `clean_structure`
5. Verify via `read_document(file_path, mode="metadata")`

---

## Antigravity IDE Configuration

### Option A: Workspace MCP Config (Automatic)
The project includes `.agents/mcp_config.json`. When opening this repository in Antigravity IDE, the server is automatically discovered:

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

> Replace `<repo-path>` with the absolute path to where you cloned this repository.

---

## Testing & Verification

Run the comprehensive unit test suite inside `.venv`:

```powershell
.venv\Scripts\python.exe -m pytest tests/ -v
```

Run the end-to-end demo:

```powershell
.venv\Scripts\python.exe demo.py
```
