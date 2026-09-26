"""FastMCP Word DOCX Server.

Exposes 4 consolidated endpoints with 0-based indexing contracts and low-level XML mutations
for token-efficient Word document manipulation in Antigravity IDE.
"""

from typing import Any, Dict, List, Literal, Optional, Union
from fastmcp import FastMCP
from src.oxml_engine import DocxEngine

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
    return DocxEngine.read_document(file_path, mode=mode, table_index=table_index)


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
    return DocxEngine.compose_document(file_path, actions=actions, create_if_missing=create_if_missing)


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
    return DocxEngine.mutate_element(
        file_path,
        target_index=target_index,
        action=action,
        content=content,
    )


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
    return DocxEngine.find_and_replace(
        file_path,
        find_text=find_text,
        replace_text=replace_text,
        match_case=match_case,
        scope=scope,
    )


if __name__ == "__main__":
    mcp.run()
