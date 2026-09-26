# Implementation Plan: Word DOCX FastMCP Server for Antigravity IDE (Refined)

## 1. Objective
Build an optimized, token-efficient Model Context Protocol (MCP) server using **FastMCP** and **python-docx**, running in a pinned Python 3.11 virtual environment (`.venv`). The server exposes a lean, consolidated API (4 endpoints instead of 14), uses low-level OpenXML (`oxml`/`lxml`) mutations for precise, in-place paragraph insertions and deletions, and strictly enforces **0-based indexing** across all inspection outputs and mutation schemas to prevent off-by-one LLM hallucinations.

---

## 2. Architecture & File Structure

```text
MsWord-ai-integration/
├── .venv/                              # Isolated Python 3.11 virtual environment
├── .agent-info/                        # Agent summaries, research, and planning notes
├── .agents/
│   ├── rules/                          # Workspace rules (problem-resolve, info-dir)
│   └── plugins/
│       └── word-docx/
│           ├── plugin.json             # Workspace plugin manifest
│           └── mcp_config.json         # Workspace MCP config for Antigravity IDE
├── src/
│   ├── __init__.py
│   ├── server.py                       # FastMCP server & 4 consolidated tool definitions
│   └── oxml_engine.py                  # Core python-docx + low-level XML mutation engine
├── tests/
│   ├── __init__.py
│   └── test_mutations.py               # Unit tests verifying XML insertions, deletions & batching
├── .python-version                     # Pinned to 3.11
├── requirements.txt                    # Pinned dependencies
└── README.md                           # Documentation & tool reference
```

---

## 3. Environment & Runtime Pinning

* **Python Version**: Pinned to **Python 3.11** (`py -3.11 -m venv .venv`).
  * *Reason*: `python-docx` relies on `lxml`, which requires pre-built binary wheels on Windows. Python 3.11 guarantees mature, stable wheels and eliminates C-compilation failures that occur on bleeding-edge Python 3.14.
* **Interpreter Path**:
  * `c:/Users/levine/Documents/Github/MsWord-ai-integration/.venv/Scripts/python.exe`
* **Dependencies**:
  * `fastmcp>=0.1.0`
  * `python-docx>=1.1.0`
  * `pydantic>=2.0.0`
  * `lxml>=4.9.0`

---

## 4. Consolidated Tool API (4 Tools Total) & 0-Based Indexing Contract

All elements (paragraphs, tables, rows, columns) are strictly addressed via **0-based indexing** in both outputs and tool arguments:

### 1. `read_document`
Consolidates all document inspection, outline extraction, and table viewing into one endpoint.
* **Signature**:
  ```python
  read_document(
      file_path: str,
      mode: Literal["outline", "full", "table", "metadata"] = "full",
      table_index: Optional[int] = None
  ) -> str | dict
  ```
* **Modes**:
  * `"outline"`: Headings hierarchy and element summary with explicit 0-based IDs (e.g. `[p:0] (Heading 1) Title`, `[p:1] (Normal) Body text...`, `[table:0] 3x4 grid`).
  * `"full"`: Structured Markdown representation of the entire document with 0-based indices prefixed on every element.
  * `"table"`: Extracts specific table at 0-based `table_index` as structured Markdown/JSON.
  * `"metadata"`: Core OpenXML properties (author, title, modified, revision).

---

### 2. `compose_document`
Batch content writer that appends multiple elements in a single round-trip, avoiding multiple sequential tool calls.
* **Signature**:
  ```python
  compose_document(
      file_path: str,
      actions: list[dict],
      create_if_missing: bool = True
  ) -> str
  ```
* **Action Types**:
  * `{"type": "heading", "text": "...", "level": 1}`
  * `{"type": "paragraph", "text": "...", "style": "Normal", "bold": False, "italic": False}`
  * `{"type": "list_item", "text": "...", "style": "List Bullet"}`
  * `{"type": "table", "headers": ["Col 0", "Col 1"], "rows": [["A", "B"]]}`
  * `{"type": "page_break"}`

---

### 3. `mutate_element` (Low-Level XML Engine)
Performs surgical, targeted insertions, deletions, and replacements at a specified **0-based paragraph index** using direct `oxml` / `lxml` node manipulation.
* **Signature**:
  ```python
  mutate_element(
      file_path: str,
      target_index: int,  # 0-based index matching read_document [p:N]
      action: Literal["insert_before", "insert_after", "delete", "replace"],
      content: Optional[dict] = None
  ) -> str
  ```
* **XML Implementation**:
  * `insert_before`: `target_p._p.addprevious(new_p._p)`
  * `insert_after`: `target_p._p.addnext(new_p._p)`
  * `delete`: `target_p._p.getparent().remove(target_p._p)`
  * `replace`: In-place text/run replacement or node substitution.

---

### 4. `find_and_replace`
Global or scoped text search and replacement across the document.
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
* **Capabilities**: Preserves run formatting where possible, reports match counts and modified locations.

---

## 5. Antigravity IDE Integration Configuration

### A. Workspace Plugin (`.agents/plugins/word-docx/`)
Antigravity automatically discovers and activates this server when the workspace is loaded:

**`plugin.json`**:
```json
{
  "name": "word-docx"
}
```

**`mcp_config.json`**:
```json
{
  "mcpServers": {
    "word-docx": {
      "command": "c:/Users/levine/Documents/Github/MsWord-ai-integration/.venv/Scripts/python.exe",
      "args": ["c:/Users/levine/Documents/Github/MsWord-ai-integration/src/server.py"],
      "env": {
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

### B. Global Option (`~/.gemini/config/mcp_config.json`)
If desired for all projects:
```json
{
  "mcpServers": {
    "word-docx": {
      "command": "c:/Users/levine/Documents/Github/MsWord-ai-integration/.venv/Scripts/python.exe",
      "args": ["c:/Users/levine/Documents/Github/MsWord-ai-integration/src/server.py"]
    }
  }
}
```

---

## 6. Execution Steps

1. **Step 1: Pinned Virtual Environment**:
   - Execute `py -3.11 -m venv .venv`.
   - Install `fastmcp`, `python-docx`, `pydantic`, and `lxml`.
   - Freeze dependencies into `requirements.txt` and create `.python-version`.

2. **Step 2: Low-Level XML & DOCX Engine (`src/oxml_engine.py`)**:
   - Implement `DocxEngine` with XML-level node manipulation:
     - `addprevious()`, `addnext()`, `parent.remove()` for paragraphs.
     - Outline generator with strict **0-based paragraph and table indexing** (`[p:0]`, `[table:0]`) for mutation targeting.
     - Run-aware find-and-replace algorithm.

3. **Step 3: FastMCP Server (`src/server.py`)**:
   - Instantiate `FastMCP("word-docx")`.
   - Expose the 4 consolidated endpoints with Pydantic type validation and docstrings clearly emphasizing 0-based indices.

4. **Step 4: Automated Testing (`tests/test_mutations.py`)**:
   - Test sequential batch composition.
   - Test XML `insert_before`, `insert_after`, and `delete` at specific 0-based paragraph indices.
   - Validate outline 0-based index tracking and mutation round-trips.

5. **Step 5: Antigravity IDE Plugin Setup**:
   - Write `.agents/plugins/word-docx/plugin.json` and `.agents/plugins/word-docx/mcp_config.json`.

---

## 7. Architectural Thoughts & Analysis

### 1. 0-Based Indexing Contract (Preventing Off-By-One Hallucinations)
* LLMs trained heavily on Python code naturally default to 0-based indexing, whereas human users or document outlines often intuitively expect 1-based indexing. Mixing these conventions creates off-by-one errors where an agent targets paragraph 1 intending to delete the first paragraph, but deletes the second.
* By enforcing a rigid contract—where `read_document` outputs explicit identifiers like `[p:0]`, `[p:1]`, `[table:0]` and the schema for `mutate_element(target_index)` explicitly requires the same `0-based` index—the agent can copy-paste the exact integer directly from inspection into mutation with zero mental arithmetic.

### 2. Token Bandwidth & Context Preservation
* In Antigravity IDE and modern LLMs, every registered MCP tool exports its full JSON-schema (names, descriptions, parameter types, nested objects) into the model's system prompt on **every single conversational turn**.
* Exposing 14 fine-grained tools (`add_heading`, `add_paragraph`, `add_table`, `add_row`, `get_heading`, etc.) consumes ~1,800–2,500 prompt tokens per turn.
* Consolidating into **4 polymorphic endpoints** (`read_document`, `compose_document`, `mutate_element`, `find_and_replace`) reduces this footprint by **~65–70%** (~600–800 tokens), preserving maximum context window for the user's actual document text and prompt reasoning.

### 3. Multi-Action Batching vs. Round-Trip Latency
* A common pitfall in AI document generation is executing 10 tool calls to write a 10-paragraph document (each requiring LLM generation -> tool dispatch -> execution -> response injection -> next generation).
* `compose_document` allows the LLM to output a single array of typed operations (`heading`, `paragraph`, `list_item`, `table`) that the server writes sequentially in **one single round-trip**.

### 4. The Necessity of Direct XML Mutations (`oxml`)
* Standard `python-docx` is fundamentally **append-only** (`document.add_paragraph()`, `document.add_heading()`).
* While `python-docx` provides an `.insert_paragraph_before()` method, it has significant limitations:
  1. It only inserts before, not after.
  2. There is **no native `delete_paragraph()`** in `python-docx`'s public API. The documented workaround is to clear the text, which leaves empty XML nodes (`<w:p>`) that corrupt paragraph spacing.
  3. Inserting tables between existing paragraphs is completely unsupported in public `python-docx`.
* By operating on the underlying `lxml.etree` elements (`_p.addprevious()`, `_p.addnext()`, and `_p.getparent().remove(_p)`), we achieve surgical, zero-artifact mutations that keep the OpenXML schema clean and valid.

### 5. Runtime Stability (Python 3.11 vs 3.14)
* `python-docx` relies on `lxml`, which is a C-extension wrapping `libxml2` and `libxslt`.
* On Windows, Python 3.14 often requires Microsoft Visual C++ 14.0+ Build Tools and manual compilation because pre-built `.whl` binaries are not yet published for new minor Python releases.
* Pinning to **Python 3.11** ensures instant, pip-wheel installation without native compilation dependencies or compiler mismatch errors.
