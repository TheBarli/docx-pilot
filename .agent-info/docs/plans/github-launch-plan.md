# Change Documentation — MsWord-AI-Integration GitHub Launch

**Date:** 2026-10-28
**Branch:** `main`
**Remote:** `https://github.com/TheBarli/MsWord-ai-integration.git`
**Commit scope:** Staged, pending push

---

## Overview

This document covers all changes made during the GitHub launch cleanup session. The session had two tracks running in parallel:

1. **Feature expansion** — the MCP server grew from 4 → 8 tools, with new engine methods added to support document styling presets.
2. **Workspace hygiene** — scratch files, trial documents, one-off scripts, and internal agent notes were removed or gitignored; README and plugin config were updated to match reality.

---

## 1. `src/server.py` — 4 → 8 MCP Tools

### Summary
The server docstring and tool count were updated from **4 consolidated endpoints** to **8**. Four new `@mcp.tool()` endpoints were added, each delegating to a corresponding `DocxEngine` static method.

### New Tools

#### `set_page_layout`
```python
set_page_layout(
    file_path: str,
    top_cm: float = 2.54,
    bottom_cm: float = 2.54,
    left_cm: float = 2.54,
    right_cm: float = 2.54,
    page_size: Literal["letter", "a4", "legal"] = "letter",
    orientation: Literal["portrait", "landscape"] = "portrait",
) -> Dict[str, Any]
```
Modifies margins and paper dimensions across **all** `<w:sectPr>` elements in `document.xml` (not just the root body section). Operates via direct ZIP/XML manipulation, bypassing `python-docx` to ensure section-level margins are uniformly applied.

---

#### `update_styles`
```python
update_styles(
    file_path: str,
    styles_config: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]
```
Configures font family, font size, bold, line spacing, space before/after, `keepNext`, and color in `word/styles.xml` for named styles (e.g. `Normal`, `Heading 1`, `Heading 2`). All child element reordering follows the ECMA-376 sequence spec to prevent schema validation errors.

---

#### `format_table_grid`
```python
format_table_grid(
    file_path: str,
    table_indices: Optional[List[int]] = None,
    border_sz: int = 4,
    border_color: str = "000000",
    header_shading: Optional[str] = "F2F2F2",
    cant_split: bool = True,
    repeat_header: bool = True,
    font_name: Optional[str] = "Times New Roman",
    font_size_pt: Optional[float] = 11.0,
) -> Dict[str, Any]
```
Normalizes table borders, row pagination properties (`<w:cantSplit/>`, `<w:tblHeader/>`), header shading, cell padding, and optional table-level font. Accepts a list of 0-based `table_indices`; `None` targets all tables in the document.

---

#### `clean_structure`
```python
clean_structure(
    file_path: str,
    remove_empty_paragraphs: bool = True,
    remove_empty_tables: bool = True,
    extract_single_image_tables: bool = True,
    max_consecutive_breaks: int = 1,
    normalize_all_colors_black: bool = False,
) -> Dict[str, Any]
```
Automated hygiene pass:
- Strips trailing empty `<w:p>` nodes (no text, drawings, or page breaks).
- Removes tables containing 0 text and 0 images/drawings.
- Extracts images from single-image container tables and removes the empty table shell.
- Purges rogue consecutive page breaks above the configured maximum.
- Optionally forces all font color tags in `document.xml` to `#000000`.

### Other Changes in `server.py`
- `os.chdir(PROJECT_ROOT)` added at startup to anchor the process working directory, fixing relative path resolution when the server is spawned by the MCP host from an arbitrary working directory.

---

## 2. `src/oxml_engine.py` — Engine Expansion

### New Imports
- `import io` — for in-memory ZIP read/write during page layout mutation
- `from lxml import etree` — direct lxml access for low-level XML manipulation

### New Module-Level Constants
```python
PROJECT_ROOT = Path(__file__).resolve().parent.parent
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
```

### ECMA-376 Schema Sequence Orders
Three ordered lists were added to drive safe XML child ordering:
- `STYLE_ORDER` — valid child element sequence for `<w:style>`
- `PPR_ORDER` — valid child element sequence for `<w:pPr>`
- `RPR_ORDER` — valid child element sequence for `<w:rPr>`

These are consumed by the new `reorder_children()` utility to prevent Word's "repaired content" warnings on open.

### New Utility Function
```python
def reorder_children(parent: etree._Element, order_list: List[str]) -> None
```
Sorts an XML element's children according to the ECMA-376 sequence specification, ensuring schema compliance after programmatic mutations.

### Updated: `normalize_path()`
**Before:** Simple `Path(file_path).expanduser().resolve()` — failed silently for relative paths when the server's CWD differed from the workspace root.

**After:** Multi-candidate resolution:
1. Absolute path that exists → return directly.
2. `CWD / relative_path` → return if it exists.
3. `PROJECT_ROOT / relative_path` → fallback (ensures paths like `"documents/Tugas 4.docx"` resolve correctly regardless of launch CWD).

### New `DocxEngine` Static Methods
Four static methods added to the `DocxEngine` class, backing the four new MCP tools listed in §1:

| Method | Key Technique |
|--------|---------------|
| `set_page_layout` | In-memory ZIP I/O + lxml `<w:sectPr>` mutation |
| `update_styles` | `word/styles.xml` parsing, `reorder_children` for ECMA-376 compliance |
| `format_table_grid` | Per-row/cell lxml mutations for borders, shading, pagination flags |
| `clean_structure` | DOM traversal; image extraction from single-image tables |

---

## 3. New Agent Customizations

### `.agents/presets/` (new directory, 3 files)

These YAML presets define reusable formatting profiles consumed by the `document-styling` rule. Each preset maps directly to the 4 new MCP tool signatures.

#### `academic-id.yaml`
Indonesian academic report standard (Dikti/Kemendikbud style):

| Section | Key Values |
|---------|-----------|
| **layout** | Letter, portrait; margins 4/3/4/3 cm (top/bottom/left/right) |
| **Normal** | Times New Roman 12pt, 1.5× line spacing, justified, 0 space after |
| **Heading 1/2** | Times New Roman 12pt bold, 1.5×, left-aligned, `keepNext: true` |
| **table** | 0.5 pt black borders, `F2F2F2` header shading, Times New Roman 11pt |
| **hygiene** | All flags on, `normalize_all_colors_black: true` |

#### `academic-id-advanced.yaml`
Extended variant with additional heading levels (H3–H5), `List Paragraph` style configuration, and per-style `space_before_pt` overrides for stricter hierarchy control.

#### `clean-uniform.yaml`
Minimal normalization preset: uniform 2.54 cm margins, Calibri body font, no color normalization. Intended for general-purpose cleanup rather than academic compliance.

---

### `.agents/rules/document-styling.md` (new file)

Conditional agent rule (triggers when the user references a formatting preset by name). Instructs the agent to:

1. Load `.agents/presets/<name>.yaml`
2. Run the styling pipeline in sequence:
   ```
   set_page_layout  →  update_styles  →  format_table_grid  →  clean_structure
   ```
3. Verify with `read_document(mode="metadata")`

---

## 4. Workspace Hygiene

### `.gitignore` — Expanded

Added the following ignore patterns:

```gitignore
# Root-level one-off pipeline scripts
insert_fixed_into_tugas4.py
normalize_academic_complete.py
populate_fixed_complete.py
post_mcp_adjust.py
snapshot_cover.py
uniformize_styling.py

# Root scratch directory
scratch/

# Agent internal scratch & ephemeral notes
.agent-info/scratch/
.agent-info/cover_page_snapshot.xml
.agent-info/merge_report.py
.agent-info/mcp_decoupled_architecture_plan.md
.agent-info/mcp_formatting_plan.md

# PDFs
*.pdf
documents/*
!documents/.gitkeep
```

### Deleted from Tracking
- `.agents/plugins/word-docx/mcp_config.json`
- `.agents/plugins/word-docx/plugin.json`

These plugin manifest files were removed. The workspace MCP server is now registered exclusively via `.agents/mcp_config.json`.

### Deleted from Disk
All trial `.docx` and `.pdf` files under `documents/` were removed. Only `documents/.gitkeep` remains to preserve the directory for user-placed files.

---

## 5. `README.md` — Updates

| Section | Change |
|---------|--------|
| **File Structure tree** | Removed `.venv/` and `plugins/` subtree; added `.agents/presets/` and updated rules list; added `documents/` with note |
| **Option A config** | Updated from "Workspace Plugin" (referencing deleted plugin files) → "Workspace MCP Config" (referencing `.agents/mcp_config.json`) |
| **Option B config** | Hardcoded absolute path `c:/Users/levine/...` replaced with `<repo-path>` placeholder + explanatory note |

---

## 6. Breaking Changes

- **Plugin manifests removed:** `.agents/plugins/word-docx/mcp_config.json` and `plugin.json` no longer exist. Anyone using that registration path must switch to `.agents/mcp_config.json` (see README Option A).
- **`normalize_path()` resolution order changed:** relative paths now resolve against CWD, then `PROJECT_ROOT`. Since `server.py` calls `os.chdir(PROJECT_ROOT)` at startup, CWD effectively equals `PROJECT_ROOT` for the server process.

---

## 7. Pre-Push Plan

Work top to bottom. Phases 1-2 are blockers for the push.

### Phase 1: Repo hygiene (blocker, ~15 min)

- [ ] **Fix `.gitignore`**: use `documents/*` (not `documents/`) so `!documents/.gitkeep` actually works.
- [ ] **Untrack already-committed files** that are now ignored:
  ```bash
  git ls-files scratch .agent-info documents *.pdf \
    insert_fixed_into_tugas4.py normalize_academic_complete.py \
    populate_fixed_complete.py post_mcp_adjust.py \
    snapshot_cover.py uniformize_styling.py
  git rm -r --cached <each path listed above>
  ```
- [ ] **Check history for leaks** (personal coursework, `c:/Users/levine/...` path):
  ```bash
  git log --all --oneline -- documents/ "*.docx" "*.pdf"
  git log --all -S"c:/Users/levine" --oneline
  ```
  If anything shows up and you care, rewrite history (`git filter-repo`) **before** the first public push.
- [ ] **Verify the staged set**: `git status` and `git diff --cached --stat` match this document.

### Phase 2: Tests for new engine methods (blocker, ~1-2 hrs)

Add to `tests/test_mutations.py`. Use a small fixture `.docx` copied to a temp dir; never mutate the fixture.

- [ ] `set_page_layout`: margins and page size applied to **every** `<w:sectPr>` in a multi-section doc; landscape swaps width/height.
- [ ] `clean_structure`: removes trailing empty paragraphs; removes empty tables; extracts image from a single-image table; caps consecutive page breaks; leaves colors alone when `normalize_all_colors_black=False`.
- [ ] `update_styles`: output child order matches `STYLE_ORDER` / `PPR_ORDER` / `RPR_ORDER`.
- [ ] `format_table_grid`: `cantSplit` and `tblHeader` present; `table_indices=None` hits all tables, a list hits only those.
- [ ] `normalize_path`: absolute, CWD-relative, and `PROJECT_ROOT`-relative cases.
- [ ] **Round-trip check**: each output reopens in `python-docx` without error (cheap proxy for Word's "repaired content" warning).

### Phase 3: README (blocker, ~30 min)

- [ ] Update the tool API section from 4 to 8 tools; document signatures and defaults for `set_page_layout`, `update_styles`, `format_table_grid`, `clean_structure`.
- [ ] Add a "Presets" section: list the 3 presets, the pipeline order, and the `read_document(mode="metadata")` verification step.
- [ ] Add a short "Migrating from the plugin config" note (see Breaking Changes).
- [ ] Warn that `clean_structure` edits files in place and `normalize_all_colors_black` flattens intentional colors (hyperlinks, highlights).

### Phase 4: Quick safety and polish (recommended, not blocking)

- [ ] Confirm `academic-id.yaml` page size: Letter vs. A4 (Indonesian academic norm is usually A4). Fix or add a comment explaining the choice.
- [ ] Add a comment above `os.chdir(PROJECT_ROOT)` explaining why it exists.
- [ ] Add a backup (`<name>.bak.docx`) or `dry_run` option to `clean_structure`.
- [ ] Validate preset YAML keys on load so typos fail loudly.

### Phase 5: Commit and push

- [ ] Run the full test suite; confirm green.
- [ ] **Commit**: `chore: clean workspace for GitHub launch` (consider splitting the 4-tool feature work into its own `feat:` commit).
- [ ] **Push** to `origin main`: requires explicit confirmation (HITL gate).
- [ ] After push, open the repo on GitHub in a private window and confirm no `documents/` contents, scratch files, or local paths are visible.