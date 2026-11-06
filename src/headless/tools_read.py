"""Document reading and structural inspection mixin for headless engine."""

from typing import Any, Dict, List, Literal, Optional, Union

from src.core.constants import OUTLINE_LINE_LIMIT
from src.headless.xml_helpers import get_or_create_document


class ReadTools:
    """Mixin providing read_document for DocxEngine."""

    @staticmethod
    def read_document(
        file_path: str,
        mode: Literal["outline", "full", "table", "metadata"] = "full",
        table_index: Optional[int] = None,
    ) -> Union[str, Dict[str, Any]]:
        """Read document using specified mode with strict 0-based indexing.

        Modes:
            outline: High-level hierarchy of headings, paragraph counts, and tables.
            full: Complete structured Markdown dump with [p:N] and [table:N] annotations.
            table: Extracts a specific table (0-based table_index) as Markdown and structured rows.
            metadata: Core OpenXML metadata properties.
        """
        doc, path = get_or_create_document(file_path, create_if_missing=False)

        if mode == "metadata":
            props = doc.core_properties
            return {
                "file_path": str(path),
                "title": props.title or "",
                "author": props.author or "",
                "subject": props.subject or "",
                "keywords": props.keywords or "",
                "category": props.category or "",
                "comments": props.comments or "",
                "created": str(props.created) if props.created else "",
                "modified": str(props.modified) if props.modified else "",
                "last_modified_by": props.last_modified_by or "",
                "revision": props.revision or 0,
                "total_paragraphs": len(doc.paragraphs),
                "total_tables": len(doc.tables),
            }

        if mode == "outline":
            lines = [
                f"# Outline: {path.name}",
                f"- **Total Paragraphs**: {len(doc.paragraphs)} (indices: [p:0] to [p:{max(0, len(doc.paragraphs)-1)}])",
                f"- **Total Tables**: {len(doc.tables)}",
                "",
                "### Structural Elements (0-based indices):",
            ]
            
            # Walk document body elements in document order
            p_idx = 0
            tbl_idx = 0
            for child in doc.element.body:
                if child.tag.endswith("p"):
                    if p_idx < len(doc.paragraphs):
                        p = doc.paragraphs[p_idx]
                        text_snippet = p.text.strip()
                        style_name = p.style.name if p.style else "Normal"
                        is_heading = style_name.lower().startswith("heading") or style_name.lower() == "title"
                        if is_heading or (len(lines) < OUTLINE_LINE_LIMIT and text_snippet):
                            preview = (text_snippet[:80] + "...") if len(text_snippet) > 80 else text_snippet
                            preview_text = f": \"{preview}\"" if preview else " (empty)"
                            lines.append(f"- `[p:{p_idx}]` **({style_name})**{preview_text}")
                    p_idx += 1
                elif child.tag.endswith("tbl"):
                    if tbl_idx < len(doc.tables):
                        tbl = doc.tables[tbl_idx]
                        row_cnt = len(tbl.rows)
                        col_cnt = len(tbl.columns) if row_cnt > 0 else 0
                        # Snippet of first row
                        header_preview = [c.text.strip() for c in tbl.rows[0].cells[:4]] if row_cnt > 0 else []
                        header_str = " | ".join(header_preview)
                        lines.append(f"- `[table:{tbl_idx}]` **Table ({row_cnt} rows x {col_cnt} cols)**: [{header_str}]")
                    tbl_idx += 1

            return "\n".join(lines)

        if mode == "table":
            if table_index is None:
                raise ValueError("table_index (0-based) is required when mode='table'")
            if table_index < 0 or table_index >= len(doc.tables):
                raise IndexError(
                    f"Table index {table_index} is out of range. Document has {len(doc.tables)} tables (indices 0 to {len(doc.tables)-1})."
                )
            tbl = doc.tables[table_index]
            matrix: List[List[str]] = []
            for row in tbl.rows:
                matrix.append([cell.text.strip() for cell in row.cells])

            md_lines = [f"### Table `[table:{table_index}]` ({len(tbl.rows)} rows x {len(tbl.columns)} cols)", ""]
            if matrix:
                header = matrix[0]
                md_lines.append("| " + " | ".join(header) + " |")
                md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")
                for row in matrix[1:]:
                    md_lines.append("| " + " | ".join(row) + " |")
            return {
                "table_index": table_index,
                "rows_count": len(tbl.rows),
                "columns_count": len(tbl.columns),
                "matrix": matrix,
                "markdown": "\n".join(md_lines),
            }

        # Default mode: "full"
        lines = [f"# Document Content: {path.name}", ""]
        p_idx = 0
        tbl_idx = 0
        for child in doc.element.body:
            if child.tag.endswith("p"):
                if p_idx < len(doc.paragraphs):
                    p = doc.paragraphs[p_idx]
                    style_name = p.style.name if p.style else "Normal"
                    text = p.text
                    lines.append(f"`[p:{p_idx}]` ({style_name}): {text}")
                p_idx += 1
            elif child.tag.endswith("tbl"):
                if tbl_idx < len(doc.tables):
                    tbl = doc.tables[tbl_idx]
                    lines.append(f"\n`[table:{tbl_idx}]` Table ({len(tbl.rows)}x{len(tbl.columns)}):")
                    for row_i, row in enumerate(tbl.rows):
                        row_vals = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                        lines.append(f"  row {row_i}: | " + " | ".join(row_vals) + " |")
                    lines.append("")
                tbl_idx += 1

        return "\n".join(lines)

