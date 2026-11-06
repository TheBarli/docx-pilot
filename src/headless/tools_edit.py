"""Document content composition, XML mutation, search/replace, and debloating mixin."""

from __future__ import annotations

import io
import os
from pathlib import Path
import shutil
import tempfile
from typing import Any, Dict, List, Literal, Optional, Tuple, Union
import xml.etree.ElementTree as ET
import zipfile

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree

from src.core.paths import normalize_path
from src.headless.xml_helpers import get_or_create_document


class EditTools:
    """Mixin providing compose, mutate, find/replace, and debloat for DocxEngine."""

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
            start_p = max(0, target_index - 2)
            end_p = min(len(doc.paragraphs), target_index + 3)
            outline = [
                f"`[p:{i}]` ({doc.paragraphs[i].style.name if doc.paragraphs[i].style else 'Normal'}): "
                f"{(doc.paragraphs[i].text[:60] + '...') if len(doc.paragraphs[i].text) > 60 else doc.paragraphs[i].text}"
                for i in range(start_p, end_p)
            ]
            return {
                "status": "success",
                "action": "delete",
                "target_index": target_index,
                "file_path": str(path),
                "remaining_paragraphs": len(doc.paragraphs),
                "local_outline": outline,
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
            start_p = max(0, target_index - 2)
            end_p = min(len(doc.paragraphs), target_index + 3)
            outline = [
                f"`[p:{i}]` ({doc.paragraphs[i].style.name if doc.paragraphs[i].style else 'Normal'}): "
                f"{(doc.paragraphs[i].text[:60] + '...') if len(doc.paragraphs[i].text) > 60 else doc.paragraphs[i].text}"
                for i in range(start_p, end_p)
            ]
            return {
                "status": "success",
                "action": "replace",
                "target_index": target_index,
                "file_path": str(path),
                "updated_text": new_text,
                "local_outline": outline,
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
            start_p = max(0, target_index - 2)
            end_p = min(len(doc.paragraphs), target_index + 3)
            outline = [
                f"`[p:{i}]` ({doc.paragraphs[i].style.name if doc.paragraphs[i].style else 'Normal'}): "
                f"{(doc.paragraphs[i].text[:60] + '...') if len(doc.paragraphs[i].text) > 60 else doc.paragraphs[i].text}"
                for i in range(start_p, end_p)
            ]
            return {
                "status": "success",
                "action": action,
                "target_index": target_index,
                "file_path": str(path),
                "total_paragraphs": len(doc.paragraphs),
                "total_tables": len(doc.tables),
                "local_outline": outline,
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
        if len(find_text) > 255:
            raise ValueError("find_text exceeds Word 255 character limit.")

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

    @staticmethod
    def debloat_document(file_path: str) -> Dict[str, Any]:
        """Strip embedded font bloat, remove font embedding flags, and recompress media.

        Matches lean academic document standards by removing unneeded embedded system
        fonts (which can add 15-25MB of bloat) and enforcing ZIP_DEFLATED on all entries.
        """
        path = normalize_path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        size_before = os.path.getsize(path)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(str(path), "r") as zin, zipfile.ZipFile(
                str(tmp_path), "w", compression=zipfile.ZIP_DEFLATED
            ) as zout:
                for item in zin.infolist():
                    if item.filename.startswith("word/fonts/") or item.filename == "word/_rels/fontTable.xml.rels":
                        continue

                    content = zin.read(item.filename)

                    if item.filename == "word/settings.xml":
                        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                        ET.register_namespace("w", ns["w"])
                        root = ET.fromstring(content)
                        for tag in ["w:embedTrueTypeFonts", "w:doNotEmbedSystemFonts", "w:saveSubsetFonts"]:
                            el = root.find(tag, ns)
                            if el is not None:
                                root.remove(el)
                        content = ET.tostring(root, xml_declaration=True, encoding="utf-8")

                    if item.filename == "word/fontTable.xml":
                        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                        ET.register_namespace("w", ns["w"])
                        root = ET.fromstring(content)
                        for font in root.findall("w:font", ns):
                            for embed_tag in ["w:embedRegular", "w:embedBold", "w:embedItalic", "w:embedBoldItalic"]:
                                em = font.find(embed_tag, ns)
                                if em is not None:
                                    font.remove(em)
                        content = ET.tostring(root, xml_declaration=True, encoding="utf-8")

                    new_item = zipfile.ZipInfo(item.filename)
                    new_item.compress_type = zipfile.ZIP_DEFLATED
                    new_item.external_attr = item.external_attr
                    zout.writestr(new_item, content)

            # Integrity check
            test_doc = docx.Document(str(tmp_path))
            p_count = len(test_doc.paragraphs)
            t_count = len(test_doc.tables)

            shutil.move(str(tmp_path), str(path))
            size_after = os.path.getsize(path)
            saved = size_before - size_after

            return {
                "status": "success",
                "file_path": str(path),
                "size_before_bytes": size_before,
                "size_after_bytes": size_after,
                "space_saved_bytes": saved,
                "space_saved_mb": round(saved / (1024 * 1024), 2),
                "paragraphs_count": p_count,
                "tables_count": t_count,
            }
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

