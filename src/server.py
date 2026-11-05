"""FastMCP Word DOCX Server.

Exposes 8 consolidated endpoints with 0-based indexing contracts, low-level XML mutations,
and automatic detection between Live Microsoft Word (COM) and Headless (python-docx/lxml) backends
for token-efficient Word document manipulation in Antigravity IDE.
"""

import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Union
from fastmcp import FastMCP

# Ensure project root is in sys.path and process working directory is anchored
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
os.chdir(PROJECT_ROOT)

try:
    from src.oxml_engine import DocxEngine
except ImportError:
    from oxml_engine import DocxEngine

try:
    from src.live_engine import (
        LiveWordEngine,
        pick_backend,
        LiveWordError,
        ERROR_MESSAGES,
    )
except ImportError:
    from live_engine import (
        LiveWordEngine,
        pick_backend,
        LiveWordError,
        ERROR_MESSAGES,
    )

# Initialize FastMCP Server
mcp = FastMCP("word-docx")


@mcp.tool()
def read_document(
    file_path: str,
    mode: Literal["outline", "full", "table", "metadata"] = "full",
    table_index: Optional[int] = None,
) -> Union[str, Dict[str, Any]]:
    """Read and inspect Word document (.docx) structure and content.

    All returned element identifiers use strict 0-based indexing ([p:0], [table:0], etc.)
    to eliminate off-by-one hallucinations.

    Parameters:
        file_path: Path to the .docx file.
        mode:
            - "outline": High-level summary of headings and elements with 0-based indices.
            - "full": Complete document dump with [p:N] and [table:N] annotations on each line.
            - "table": Extracts table at 0-based table_index as Markdown and structured matrix.
            - "metadata": Core properties (author, title, revision, timestamps).
        table_index: 0-based table index (required when mode="table", e.g. 0 for first table).
    """
    backend, com_doc = pick_backend(file_path, tool_name="read_document", is_mutation=False)
    if backend == "live" and com_doc is not None:
        res = LiveWordEngine.read_document(com_doc, file_path, mode=mode, table_index=table_index)
        if isinstance(res, str):
            return f"[backend: live]\n{res}"
        if isinstance(res, dict):
            res["backend"] = "live"
            return res
        return res

    res = DocxEngine.read_document(file_path, mode=mode, table_index=table_index)
    if isinstance(res, str):
        return f"[backend: headless]\n{res}"
    if isinstance(res, dict):
        res["backend"] = "headless"
        return res
    return res


@mcp.tool()
def compose_document(
    file_path: str,
    actions: List[Dict[str, Any]],
    create_if_missing: bool = True,
) -> Dict[str, Any]:
    """Append a batch of formatted elements to a Word document in a single call.

    Minimizes round trips by writing multi-element content sequentially.

    Parameters:
        file_path: Path to the .docx file.
        actions: List of element action dictionaries. Supported types:
            - {"type": "heading", "text": "Section 1", "level": 1} (levels 1-9)
            - {"type": "paragraph", "text": "Hello world", "style": "Normal", "bold": False, "italic": False}
            - {"type": "list_item", "text": "Bullet point", "style": "List Bullet"}
            - {"type": "table", "headers": ["Col 0", "Col 1"], "rows": [["val0", "val1"]], "style": "Table Grid"}
            - {"type": "page_break"}
        create_if_missing: If True, creates the document if it doesn't already exist.
    """
    backend, com_doc = pick_backend(file_path, tool_name="compose_document", is_mutation=True)
    if backend == "live":
        raise LiveWordError(ERROR_MESSAGES["tool_not_yet_live_capable"].format(tool_name="compose_document"))

    res = DocxEngine.compose_document(file_path, actions=actions, create_if_missing=create_if_missing)
    res["backend"] = "headless"
    return res


@mcp.tool()
def mutate_element(
    file_path: str,
    target_index: int,
    action: Literal["insert_before", "insert_after", "delete", "replace"],
    content: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Perform targeted XML-level mutations (insertions, deletions, replacements) at a 0-based paragraph index.

    Uses low-level OpenXML node operations for clean manipulation without empty artifacts.

    Parameters:
        file_path: Path to the .docx file.
        target_index: 0-based index of the target paragraph (matching [p:N] from read_document).
        action:
            - "insert_before": Inserts element directly preceding [p:target_index].
            - "insert_after": Inserts element directly following [p:target_index].
            - "delete": Deletes the target paragraph completely from the XML tree.
            - "replace": Replaces the target paragraph's text/style in place.
        content: Content dictionary for insertion/replacement (e.g. {"type": "paragraph", "text": "...", "bold": True},
                 or {"type": "heading", "text": "...", "level": 2}, or {"type": "table", "headers": [...], "rows": [...]}).
    """
    backend, com_doc = pick_backend(file_path, tool_name="mutate_element", is_mutation=True)
    if backend == "live" and com_doc is not None:
        res = LiveWordEngine.mutate_element(
            com_doc,
            file_path,
            target_index=target_index,
            action=action,
            content=content,
        )
        res["backend"] = "live"
        return res

    res = DocxEngine.mutate_element(
        file_path,
        target_index=target_index,
        action=action,
        content=content,
    )
    res["backend"] = "headless"
    return res


@mcp.tool()
def find_and_replace(
    file_path: str,
    find_text: str,
    replace_text: str,
    match_case: bool = True,
    scope: Literal["all", "body", "tables", "headers_footers"] = "all",
) -> Dict[str, Any]:
    """Search and replace text across the document while preserving run formatting.

    Parameters:
        file_path: Path to the .docx file.
        find_text: Text string to find.
        replace_text: Text string to replace with.
        match_case: If True, performs case-sensitive replacement. Default is True.
        scope: Search scope ("all", "body", "tables", "headers_footers").
    """
    backend, com_doc = pick_backend(file_path, tool_name="find_and_replace", is_mutation=True)
    if backend == "live" and com_doc is not None:
        res = LiveWordEngine.find_and_replace(
            com_doc,
            file_path,
            find_text=find_text,
            replace_text=replace_text,
            match_case=match_case,
            scope=scope,
        )
        res["backend"] = "live"
        return res

    res = DocxEngine.find_and_replace(
        file_path,
        find_text=find_text,
        replace_text=replace_text,
        match_case=match_case,
        scope=scope,
    )
    res["backend"] = "headless"
    return res


@mcp.tool()
def set_page_layout(
    file_path: str,
    top_cm: float = 2.54,
    bottom_cm: float = 2.54,
    left_cm: float = 2.54,
    right_cm: float = 2.54,
    page_size: Literal["letter", "a4", "legal"] = "letter",
    orientation: Literal["portrait", "landscape"] = "portrait",
) -> Dict[str, Any]:
    """Modifies margins and dimensions across all <w:sectPr> elements in document.xml (not just the root body section).

    Parameters:
        file_path: Path to the .docx file.
        top_cm: Top margin in centimeters.
        bottom_cm: Bottom margin in centimeters.
        left_cm: Left margin in centimeters.
        right_cm: Right margin in centimeters.
        page_size: Paper size ("letter", "a4", "legal").
        orientation: Page orientation ("portrait", "landscape").
    """
    backend, com_doc = pick_backend(file_path, tool_name="set_page_layout", is_mutation=True)
    if backend == "live":
        raise LiveWordError(ERROR_MESSAGES["tool_not_yet_live_capable"].format(tool_name="set_page_layout"))

    res = DocxEngine.set_page_layout(
        file_path=file_path,
        top_cm=top_cm,
        bottom_cm=bottom_cm,
        left_cm=left_cm,
        right_cm=right_cm,
        page_size=page_size,
        orientation=orientation,
    )
    res["backend"] = "headless"
    return res


@mcp.tool()
def update_styles(
    file_path: str,
    styles_config: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """Configures font family, font size, bold, line spacing, space before/after, keepNext, and color in word/styles.xml for Normal, Heading 1, and Heading 2.

    Parameters:
        file_path: Path to the .docx file.
        styles_config: Mapping of style name (e.g., 'Normal', 'Heading 1', 'Heading 2') to style attributes:
                       'font', 'size_pt', 'bold', 'color', 'line_spacing', 'space_after_pt',
                       'space_before_pt', 'alignment', 'keep_next'.
    """
    backend, com_doc = pick_backend(file_path, tool_name="update_styles", is_mutation=True)
    if backend == "live":
        raise LiveWordError(ERROR_MESSAGES["tool_not_yet_live_capable"].format(tool_name="update_styles"))

    res = DocxEngine.update_styles(
        file_path=file_path,
        styles_config=styles_config,
    )
    res["backend"] = "headless"
    return res


@mcp.tool()
def format_table_grid(
    file_path: str,
    table_indices: Optional[List[int]] = None,
    border_sz: int = 4,
    border_color: str = "000000",
    header_shading: Optional[str] = "F2F2F2",
    cant_split: bool = True,
    repeat_header: bool = True,
    font_name: Optional[str] = "Times New Roman",
    font_size_pt: Optional[float] = 11.0,
) -> Dict[str, Any]:
    """Normalizes table borders, row pagination properties, header shading, and cell padding.

    Parameters:
        file_path: Path to the .docx file.
        table_indices: Optional list of 0-based table indices. None formats all tables.
        border_sz: Border width in 1/8 pt (4 = 0.5 pt).
        border_color: Border hex color (e.g. '000000').
        header_shading: Hex color string for header row fill (e.g. 'F2F2F2').
        cant_split: Prevent rows from splitting across page breaks (<w:cantSplit/>).
        repeat_header: Repeat header row on new pages (<w:tblHeader/>).
        font_name: Optional font name for table text (e.g. 'Times New Roman').
        font_size_pt: Optional font size for table text in points (e.g. 11.0).
    """
    backend, com_doc = pick_backend(file_path, tool_name="format_table_grid", is_mutation=True)
    if backend == "live":
        raise LiveWordError(ERROR_MESSAGES["tool_not_yet_live_capable"].format(tool_name="format_table_grid"))

    res = DocxEngine.format_table_grid(
        file_path=file_path,
        table_indices=table_indices,
        border_sz=border_sz,
        border_color=border_color,
        header_shading=header_shading,
        cant_split=cant_split,
        repeat_header=repeat_header,
        font_name=font_name,
        font_size_pt=font_size_pt,
    )
    res["backend"] = "headless"
    return res


@mcp.tool()
def clean_structure(
    file_path: str,
    remove_empty_paragraphs: bool = True,
    remove_empty_tables: bool = True,
    extract_single_image_tables: bool = True,
    max_consecutive_breaks: int = 1,
    normalize_all_colors_black: bool = False,
) -> Dict[str, Any]:
    """Automated hygiene primitive that strips trailing empty <w:p> nodes, removes empty tables with 0 text/drawings, and purges rogue page breaks.

    Parameters:
        file_path: Path to the .docx file.
        remove_empty_paragraphs: Strip paragraphs with no text, drawings, or page breaks.
        remove_empty_tables: Remove tables containing 0 text and 0 images/drawings.
        extract_single_image_tables: Move images out of single-image container tables and remove the table.
        max_consecutive_breaks: Maximum allowed consecutive page breaks (removes extras).
        normalize_all_colors_black: Force all font color tags in document.xml to #000000.
    """
    backend, com_doc = pick_backend(file_path, tool_name="clean_structure", is_mutation=True)
    if backend == "live":
        raise LiveWordError(ERROR_MESSAGES["tool_not_yet_live_capable"].format(tool_name="clean_structure"))

    res = DocxEngine.clean_structure(
        file_path=file_path,
        remove_empty_paragraphs=remove_empty_paragraphs,
        remove_empty_tables=remove_empty_tables,
        extract_single_image_tables=extract_single_image_tables,
        max_consecutive_breaks=max_consecutive_breaks,
        normalize_all_colors_black=normalize_all_colors_black,
    )
    res["backend"] = "headless"
    return res


if __name__ == "__main__":
    mcp.run()
