# Word DOCX FastMCP Server for Antigravity IDE

A high-performance, token-efficient Model Context Protocol (MCP) server for inspecting, generating, and surgically mutating Microsoft Word (`.docx`) documents directly from Antigravity IDE.

---

## Highlights

* **Consolidated 4-Tool API**: Reduces prompt token overhead by ~65-70% compared to granular 14-tool designs, reserving maximum context bandwidth for actual document text.
* **Strict 0-Based Indexing**: All paragraph (`[p:N]`), table (`[table:N]`), row, and column identifiers strictly use 0-based indexing to eliminate off-by-one hallucinations.
* **Low-Level XML Mutations**: Leverages underlying `lxml`/`oxml` element nodes (`addprevious`, `addnext`, `parent.remove`) for in-place paragraph/table insertions and clean deletions without empty-node document corruption.
* **Pinned Python 3.11 Runtime**: Guarantees pre-built C-wheels on Windows for `lxml` and `python-docx`.

---

## File Structure

```text
MsWord-ai-integration/
├── .venv/                              # Isolated Python 3.11 environment
├── .agent-info/                        # Persistent planning notes & documentation
├── .agents/
│   ├── rules/                          # Collaborative problem resolution & info rules
│   ├── mcp_config.json                 # Workspace MCP config
│   └── plugins/
│       └── word-docx/
│           ├── plugin.json             # Workspace plugin manifest
│           └── mcp_config.json         # Plugin MCP config
├── src/
│   ├── __init__.py
│   ├── server.py                       # FastMCP server entry point (4 tools)
│   └── oxml_engine.py                  # Core python-docx + XML mutation engine
├── tests/
│   ├── __init__.py
│   └── test_mutations.py               # Unit tests verifying XML mutations & indexing
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

---

## Antigravity IDE Configuration

### Option A: Workspace Plugin (Automatic)
The project includes `.agents/plugins/word-docx/mcp_config.json` and `.agents/mcp_config.json`. When opening this repository in Antigravity IDE, the server is automatically discovered.

### Option B: Global Configuration
To make this server accessible across all workspaces in Antigravity IDE, add the following entry to `~/.gemini/config/mcp_config.json`:

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

---

## Testing & Verification

Run the test suite inside `.venv`:

```powershell
.venv\Scripts\python.exe -m pytest tests/
```
