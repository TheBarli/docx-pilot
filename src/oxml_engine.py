"""Core OpenXML and python-docx Engine.

Enforces strict 0-based indexing for all document elements (paragraphs, tables, rows, cells).
Provides low-level XML mutations (addprevious, addnext, remove) for precision editing without
corrupting OpenXML schemas.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Tuple, Union
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph


def normalize_path(file_path: str) -> Path:
    """Normalize file path and resolve to absolute path."""
    path = Path(file_path).expanduser().resolve()
    return path


def get_or_create_document(file_path: str, create_if_missing: bool = False) -> Tuple[Document, Path]:
    """Load existing document or create a new one if permitted.
    
    Returns tuple of (Document, resolved_Path).
    """
    path = normalize_path(file_path)
    if not path.exists():
        if create_if_missing:
            path.parent.mkdir(parents=True, exist_ok=True)
            doc = Document()
            doc.save(str(path))
            return doc, path
        raise FileNotFoundError(f"File not found: {path}")
    return Document(str(path)), path


class DocxEngine:
    """Low-level engine for Word document inspection, composition, and XML mutation."""

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
                        if is_heading or (len(lines) < 40 and text_snippet):
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

    @staticmethod
    def compose_document(
        file_path: str,
        actions: List[Dict[str, Any]],
        create_if_missing: bool = True,
    ) -> Dict[str, Any]:
        """Execute a batch of content creation actions sequentially.

        Action format examples:
            {"type": "heading", "text": "Section 1", "level": 1}
            {"type": "paragraph", "text": "Paragraph text", "style": "Normal", "bold": False, "italic": False}
            {"type": "list_item", "text": "Item text", "style": "List Bullet"}
            {"type": "table", "headers": ["Col 0", "Col 1"], "rows": [["val0", "val1"]]}
            {"type": "page_break"}
        """
        doc, path = get_or_create_document(file_path, create_if_missing=create_if_missing)
        executed: List[str] = []

        for idx, act in enumerate(actions):
            act_type = act.get("type", "").lower()

            if act_type == "heading":
                text = act.get("text", "")
                level = int(act.get("level", 1))
                doc.add_heading(text, level=level)
                executed.append(f"#{idx} Added Heading (level {level}): \"{text[:40]}\"")

            elif act_type == "paragraph":
                text = act.get("text", "")
                style = act.get("style", "Normal")
                p = doc.add_paragraph(style=style)
                run = p.add_run(text)
                if act.get("bold"):
                    run.bold = True
                if act.get("italic"):
                    run.italic = True
                executed.append(f"#{idx} Added Paragraph ({style}): \"{text[:40]}\"")

            elif act_type == "list_item":
                text = act.get("text", "")
                style = act.get("style", "List Bullet")
                doc.add_paragraph(text, style=style)
                executed.append(f"#{idx} Added List Item ({style}): \"{text[:40]}\"")

            elif act_type == "table":
                headers: List[str] = act.get("headers", [])
                rows: List[List[str]] = act.get("rows", [])
                style = act.get("style", "Table Grid")

                total_rows = len(rows) + (1 if headers else 0)
                total_cols = len(headers) if headers else (len(rows[0]) if rows else 0)

                table = doc.add_table(rows=total_rows, cols=total_cols)
                table.style = style

                current_row = 0
                if headers:
                    hdr_cells = table.rows[0].cells
                    for c_idx, h_text in enumerate(headers):
                        hdr_cells[c_idx].text = str(h_text)
                        for run in hdr_cells[c_idx].paragraphs[0].runs:
                            run.font.bold = True
                    current_row = 1

                for row_data in rows:
                    row_cells = table.rows[current_row].cells
                    for c_idx, val in enumerate(row_data):
                        if c_idx < len(row_cells):
                            row_cells[c_idx].text = str(val)
                    current_row += 1

                executed.append(f"#{idx} Added Table ({total_rows}x{total_cols})")

            elif act_type == "page_break":
                doc.add_page_break()
                executed.append(f"#{idx} Added Page Break")
            else:
                raise ValueError(f"Unknown action type '{act_type}' at index {idx}")

        doc.save(str(path))
        return {
            "status": "success",
            "file_path": str(path),
            "actions_executed": len(executed),
            "log": executed,
            "total_paragraphs": len(doc.paragraphs),
            "total_tables": len(doc.tables),
        }

    @staticmethod
    def mutate_element(
        file_path: str,
        target_index: int,
        action: Literal["insert_before", "insert_after", "delete", "replace"],
        content: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Perform surgical XML-level mutations at a 0-based paragraph index.

        Actions:
            delete: Removes the paragraph element completely from the XML DOM.
            insert_before: Inserts a paragraph or table directly preceding [p:target_index].
            insert_after: Inserts a paragraph or table directly following [p:target_index].
            replace: Modifies the target paragraph in place.
        """
        doc, path = get_or_create_document(file_path, create_if_missing=False)

        if target_index < 0 or target_index >= len(doc.paragraphs):
            raise IndexError(
                f"Target paragraph index {target_index} is out of range. "
                f"Document has {len(doc.paragraphs)} paragraphs (valid 0-based indices: 0 to {max(0, len(doc.paragraphs)-1)})."
            )

        target_p = doc.paragraphs[target_index]
        content = content or {}
        c_type = content.get("type", "paragraph").lower()

        if action == "delete":
            # Low-level XML node removal: cleanly remove from parent body
            p_elem = target_p._p
            p_elem.getparent().remove(p_elem)
            doc.save(str(path))
            return {
                "status": "success",
                "action": "delete",
                "target_index": target_index,
                "file_path": str(path),
                "remaining_paragraphs": len(doc.paragraphs),
            }

        elif action == "replace":
            new_text = content.get("text", "")
            target_p.text = new_text
            if "style" in content:
                target_p.style = content["style"]
            if content.get("bold") or content.get("italic"):
                for run in target_p.runs:
                    if content.get("bold") is not None:
                        run.bold = bool(content["bold"])
                    if content.get("italic") is not None:
                        run.italic = bool(content["italic"])
            doc.save(str(path))
            return {
                "status": "success",
                "action": "replace",
                "target_index": target_index,
                "file_path": str(path),
                "updated_text": new_text,
            }

        elif action in ("insert_before", "insert_after"):
            # Construct the new element at the document tail first
            if c_type == "table":
                headers: List[str] = content.get("headers", [])
                rows: List[List[str]] = content.get("rows", [])
                style = content.get("style", "Table Grid")
                total_rows = len(rows) + (1 if headers else 0)
                total_cols = len(headers) if headers else (len(rows[0]) if rows else 0)

                temp_tbl = doc.add_table(rows=total_rows, cols=total_cols)
                temp_tbl.style = style

                curr = 0
                if headers:
                    for c_idx, h in enumerate(headers):
                        temp_tbl.rows[0].cells[c_idx].text = str(h)
                    curr = 1
                for r in rows:
                    for c_idx, val in enumerate(r):
                        if c_idx < total_cols:
                            temp_tbl.rows[curr].cells[c_idx].text = str(val)
                    curr += 1

                # XML mutation: move tbl._tbl node next to target_p._p
                if action == "insert_before":
                    target_p._p.addprevious(temp_tbl._tbl)
                else:
                    target_p._p.addnext(temp_tbl._tbl)

            elif c_type == "heading":
                text = content.get("text", "")
                level = int(content.get("level", 1))
                temp_p = doc.add_heading(text, level=level)
                if action == "insert_before":
                    target_p._p.addprevious(temp_p._p)
                else:
                    target_p._p.addnext(temp_p._p)

            else:  # paragraph or list_item
                text = content.get("text", "")
                style = content.get("style", "Normal")
                temp_p = doc.add_paragraph(text, style=style)
                if content.get("bold") or content.get("italic"):
                    for run in temp_p.runs:
                        if content.get("bold"):
                            run.bold = True
                        if content.get("italic"):
                            run.italic = True
                if action == "insert_before":
                    target_p._p.addprevious(temp_p._p)
                else:
                    target_p._p.addnext(temp_p._p)

            doc.save(str(path))
            return {
                "status": "success",
                "action": action,
                "target_index": target_index,
                "file_path": str(path),
                "total_paragraphs": len(doc.paragraphs),
                "total_tables": len(doc.tables),
            }

        raise ValueError(f"Unknown action: {action}")

    @staticmethod
    def find_and_replace(
        file_path: str,
        find_text: str,
        replace_text: str,
        match_case: bool = True,
        scope: Literal["all", "body", "tables", "headers_footers"] = "all",
    ) -> Dict[str, Any]:
        """Find and replace text across paragraphs, tables, headers, and footers.

        Preserves formatting wherever matches reside within individual runs, and handles
        run boundaries gracefully.
        """
        if not find_text:
            raise ValueError("find_text cannot be empty.")

        doc, path = get_or_create_document(file_path, create_if_missing=False)
        total_replacements = 0
        details: List[str] = []

        def replace_in_paragraph(p: Paragraph, loc_label: str):
            nonlocal total_replacements
            text = p.text
            if not text:
                return

            has_match = (find_text in text) if match_case else (find_text.lower() in text.lower())
            if not has_match:
                return

            # Check if match is localized within individual runs
            local_replaced = False
            for run in p.runs:
                r_text = run.text
                if not r_text:
                    continue
                match_in_run = (find_text in r_text) if match_case else (find_text.lower() in r_text.lower())
                if match_in_run:
                    if match_case:
                        run.text = r_text.replace(find_text, replace_text)
                    else:
                        import re
                        pattern = re.compile(re.escape(find_text), re.IGNORECASE)
                        run.text = pattern.sub(replace_text, r_text)
                    local_replaced = True
                    total_replacements += 1

            # If match spanned multiple runs, replace at paragraph text level
            if not local_replaced:
                if match_case:
                    count = text.count(find_text)
                    p.text = text.replace(find_text, replace_text)
                    total_replacements += count
                else:
                    import re
                    pattern = re.compile(re.escape(find_text), re.IGNORECASE)
                    matches = len(pattern.findall(text))
                    p.text = pattern.sub(replace_text, text)
                    total_replacements += matches

            details.append(loc_label)

        # 1. Body Paragraphs
        if scope in ("all", "body"):
            for p_idx, p in enumerate(doc.paragraphs):
                replace_in_paragraph(p, f"[p:{p_idx}]")

        # 2. Tables
        if scope in ("all", "tables"):
            for t_idx, tbl in enumerate(doc.tables):
                for r_idx, row in enumerate(tbl.rows):
                    for c_idx, cell in enumerate(row.cells):
                        for cp_idx, cp in enumerate(cell.paragraphs):
                            replace_in_paragraph(cp, f"[table:{t_idx} row:{r_idx} col:{c_idx}]")

        # 3. Headers & Footers
        if scope in ("all", "headers_footers"):
            for s_idx, sec in enumerate(doc.sections):
                for hp in sec.header.paragraphs:
                    replace_in_paragraph(hp, f"[section:{s_idx} header]")
                for fp in sec.footer.paragraphs:
                    replace_in_paragraph(fp, f"[section:{s_idx} footer]")

        doc.save(str(path))
        return {
            "status": "success",
            "file_path": str(path),
            "find_text": find_text,
            "replace_text": replace_text,
            "matches_replaced": total_replacements,
            "affected_locations_count": len(details),
            "sample_locations": details[:10],
        }
