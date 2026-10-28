"""Core OpenXML and python-docx Engine.

Enforces strict 0-based indexing for all document elements (paragraphs, tables, rows, cells).
Provides low-level XML mutations (addprevious, addnext, remove) for precision editing without
corrupting OpenXML schemas.
"""

from __future__ import annotations

import io
import os
from pathlib import Path
import shutil
import tempfile
from typing import Any, Dict, List, Literal, Optional, Tuple, Union
import xml.etree.ElementTree as ET
import zipfile
from lxml import etree
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

PROJECT_ROOT = Path(__file__).resolve().parent.parent
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# ECMA-376 schema sequence orders for XML validation safety
STYLE_ORDER = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
    "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
    "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
    "tblPr", "trPr", "tcPr", "tblStylePr",
]

PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi",
    "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]

RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "color", "spacing", "w", "kern", "position", "sz", "szCs",
    "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl",
    "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
]


def reorder_children(parent: etree._Element, order_list: List[str]) -> None:
    """Sort children of an XML element according to ECMA-376 sequence specification."""
    order_map = {tag: i for i, tag in enumerate(order_list)}
    children = list(parent)
    children.sort(key=lambda el: order_map.get(el.tag.split("}")[-1] if "}" in el.tag else el.tag, 999))
    for c in list(parent):
        parent.remove(c)
    for c in children:
        parent.append(c)


def normalize_path(file_path: str) -> Path:
    """Normalize file path and resolve to absolute path, checking workspace root if relative."""
    p = Path(file_path).expanduser()
    if p.is_absolute() and p.exists():
        return p.resolve()
    cwd_candidate = (Path.cwd() / p).resolve()
    if cwd_candidate.exists():
        return cwd_candidate
    root_candidate = (PROJECT_ROOT / p).resolve()
    if root_candidate.exists() or not p.is_absolute():
        return root_candidate
    return p.resolve()


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

    @staticmethod
    def set_page_layout(
        file_path: str,
        top_cm: float = 2.54,
        bottom_cm: float = 2.54,
        left_cm: float = 2.54,
        right_cm: float = 2.54,
        page_size: Literal["letter", "a4", "legal"] = "letter",
        orientation: Literal["portrait", "landscape"] = "portrait",
    ) -> Dict[str, Any]:
        """Modify margins and dimensions across ALL <w:sectPr> elements in document.xml."""
        path = normalize_path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        top_twips = str(int(round(top_cm * 567)))
        bottom_twips = str(int(round(bottom_cm * 567)))
        left_twips = str(int(round(left_cm * 567)))
        right_twips = str(int(round(right_cm * 567)))

        size_map = {
            "letter": (12240, 15840),
            "a4": (11909, 16834),
            "legal": (12240, 20160),
        }
        w_twips, h_twips = size_map.get(page_size.lower(), (12240, 15840))
        if orientation == "landscape":
            w_twips, h_twips = h_twips, w_twips

        with open(path, "rb") as f:
            src_bytes = io.BytesIO(f.read())

        with zipfile.ZipFile(src_bytes, "r") as zin:
            namelist = zin.namelist()
            doc_xml = etree.fromstring(zin.read("word/document.xml"))
            sections = doc_xml.findall(".//w:sectPr", namespaces={"w": W_NS})

            for s in sections:
                pgMar = s.find("w:pgMar", namespaces={"w": W_NS})
                if pgMar is None:
                    pgMar = etree.SubElement(s, f"{{{W_NS}}}pgMar")
                pgMar.attrib[f"{{{W_NS}}}top"] = top_twips
                pgMar.attrib[f"{{{W_NS}}}bottom"] = bottom_twips
                pgMar.attrib[f"{{{W_NS}}}left"] = left_twips
                pgMar.attrib[f"{{{W_NS}}}right"] = right_twips
                pgMar.attrib[f"{{{W_NS}}}header"] = "720"
                pgMar.attrib[f"{{{W_NS}}}footer"] = "720"

                pgSz = s.find("w:pgSz", namespaces={"w": W_NS})
                if pgSz is None:
                    pgSz = etree.SubElement(s, f"{{{W_NS}}}pgSz")
                pgSz.attrib[f"{{{W_NS}}}w"] = str(w_twips)
                pgSz.attrib[f"{{{W_NS}}}h"] = str(h_twips)
                pgSz.attrib[f"{{{W_NS}}}orient"] = orientation

            tmp_path = path.with_suffix(".tmp")
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in namelist:
                    if item == "word/document.xml":
                        zout.writestr(item, etree.tostring(doc_xml, xml_declaration=True, encoding="utf-8", standalone=True))
                    else:
                        zout.writestr(item, zin.read(item))

        if path.exists():
            os.remove(path)
        os.rename(tmp_path, path)

        return {
            "status": "success",
            "file_path": str(path),
            "sections_updated": len(sections),
            "margins_cm": {"top": top_cm, "bottom": bottom_cm, "left": left_cm, "right": right_cm},
            "page_size": page_size,
            "orientation": orientation,
        }

    @staticmethod
    def update_styles(
        file_path: str,
        styles_config: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Any]:
        """Configure font family, font size, bold, line spacing, spacing, keepNext, and color in word/styles.xml."""
        path = normalize_path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        with open(path, "rb") as f:
            src_bytes = io.BytesIO(f.read())

        with zipfile.ZipFile(src_bytes, "r") as zin:
            namelist = zin.namelist()
            styles_xml = etree.fromstring(zin.read("word/styles.xml"))

            # Schema safety: Ensure all styles have a valid <w:name>
            for s in styles_xml.findall("w:style", namespaces={"w": W_NS}):
                if s.find("w:name", namespaces={"w": W_NS}) is None:
                    sid = s.attrib.get(f"{{{W_NS}}}styleId", "Unknown")
                    new_name = etree.Element(f"{{{W_NS}}}name")
                    new_name.attrib[f"{{{W_NS}}}val"] = f"Style {sid}"
                    s.insert(0, new_name)

            updated_styles = []
            for target_name, cfg in styles_config.items():
                style_node = None
                for s in styles_xml.findall("w:style", namespaces={"w": W_NS}):
                    sid = s.attrib.get(f"{{{W_NS}}}styleId", "")
                    name_el = s.find("w:name", namespaces={"w": W_NS})
                    name_val = name_el.attrib.get(f"{{{W_NS}}}val", "") if name_el is not None else ""
                    if sid.lower() == target_name.lower().replace(" ", "") or name_val.lower() == target_name.lower():
                        style_node = s
                        break

                if style_node is None:
                    style_node = etree.SubElement(styles_xml, f"{{{W_NS}}}style")
                    style_node.attrib[f"{{{W_NS}}}type"] = "paragraph"
                    style_node.attrib[f"{{{W_NS}}}styleId"] = target_name.replace(" ", "")
                    name_el = etree.SubElement(style_node, f"{{{W_NS}}}name")
                    name_el.attrib[f"{{{W_NS}}}val"] = target_name

                # Process rPr (character formatting)
                rPr = style_node.find("w:rPr", namespaces={"w": W_NS})
                if rPr is None:
                    rPr = etree.SubElement(style_node, f"{{{W_NS}}}rPr")

                if "font" in cfg:
                    rFonts = rPr.find("w:rFonts", namespaces={"w": W_NS})
                    if rFonts is None:
                        rFonts = etree.SubElement(rPr, f"{{{W_NS}}}rFonts")
                    rFonts.attrib[f"{{{W_NS}}}ascii"] = str(cfg["font"])
                    rFonts.attrib[f"{{{W_NS}}}hAnsi"] = str(cfg["font"])
                    rFonts.attrib[f"{{{W_NS}}}cs"] = str(cfg["font"])

                if "size_pt" in cfg:
                    sz_val = str(int(round(float(cfg["size_pt"]) * 2)))
                    for sz_tag in ["sz", "szCs"]:
                        sz_el = rPr.find(f"w:{sz_tag}", namespaces={"w": W_NS})
                        if sz_el is None:
                            sz_el = etree.SubElement(rPr, f"{{{W_NS}}}{sz_tag}")
                        sz_el.attrib[f"{{{W_NS}}}val"] = sz_val

                if "bold" in cfg:
                    b_el = rPr.find("w:b", namespaces={"w": W_NS})
                    bCs_el = rPr.find("w:bCs", namespaces={"w": W_NS})
                    if cfg["bold"]:
                        if b_el is None:
                            etree.SubElement(rPr, f"{{{W_NS}}}b")
                        if bCs_el is None:
                            etree.SubElement(rPr, f"{{{W_NS}}}bCs")
                    else:
                        if b_el is not None:
                            rPr.remove(b_el)
                        if bCs_el is not None:
                            rPr.remove(bCs_el)

                if "color" in cfg:
                    c_el = rPr.find("w:color", namespaces={"w": W_NS})
                    if c_el is None:
                        c_el = etree.SubElement(rPr, f"{{{W_NS}}}color")
                    c_el.attrib[f"{{{W_NS}}}val"] = str(cfg["color"]).replace("#", "")

                # Process pPr (paragraph formatting)
                pPr = style_node.find("w:pPr", namespaces={"w": W_NS})
                if pPr is None:
                    pPr = etree.SubElement(style_node, f"{{{W_NS}}}pPr")

                if "line_spacing" in cfg or "space_after_pt" in cfg or "space_before_pt" in cfg:
                    sp_el = pPr.find("w:spacing", namespaces={"w": W_NS})
                    if sp_el is None:
                        sp_el = etree.SubElement(pPr, f"{{{W_NS}}}spacing")
                    if "line_spacing" in cfg:
                        sp_el.attrib[f"{{{W_NS}}}line"] = str(int(round(float(cfg["line_spacing"]) * 240)))
                        sp_el.attrib[f"{{{W_NS}}}lineRule"] = "auto"
                    if "space_after_pt" in cfg:
                        sp_el.attrib[f"{{{W_NS}}}after"] = str(int(round(float(cfg["space_after_pt"]) * 20)))
                    if "space_before_pt" in cfg:
                        sp_el.attrib[f"{{{W_NS}}}before"] = str(int(round(float(cfg["space_before_pt"]) * 20)))

                if "alignment" in cfg:
                    jc_el = pPr.find("w:jc", namespaces={"w": W_NS})
                    if jc_el is None:
                        jc_el = etree.SubElement(pPr, f"{{{W_NS}}}jc")
                    jc_el.attrib[f"{{{W_NS}}}val"] = str(cfg["alignment"])

                if "keep_next" in cfg:
                    kn_el = pPr.find("w:keepNext", namespaces={"w": W_NS})
                    if cfg["keep_next"]:
                        if kn_el is None:
                            etree.SubElement(pPr, f"{{{W_NS}}}keepNext")
                    else:
                        if kn_el is not None:
                            pPr.remove(kn_el)

                reorder_children(rPr, RPR_ORDER)
                reorder_children(pPr, PPR_ORDER)
                reorder_children(style_node, STYLE_ORDER)

                updated_styles.append(target_name)

            tmp_path = path.with_suffix(".tmp")
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in namelist:
                    if item == "word/styles.xml":
                        zout.writestr(item, etree.tostring(styles_xml, xml_declaration=True, encoding="utf-8", standalone=True))
                    else:
                        zout.writestr(item, zin.read(item))

        if path.exists():
            os.remove(path)
        os.rename(tmp_path, path)

        return {
            "status": "success",
            "file_path": str(path),
            "styles_updated": updated_styles,
        }

    @staticmethod
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
        """Normalize table borders, row pagination properties, header shading, and cell padding."""
        path = normalize_path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        with open(path, "rb") as f:
            src_bytes = io.BytesIO(f.read())

        with zipfile.ZipFile(src_bytes, "r") as zin:
            namelist = zin.namelist()
            doc_xml = etree.fromstring(zin.read("word/document.xml"))
            tables = doc_xml.findall(".//w:tbl", namespaces={"w": W_NS})

            target_indices = table_indices if table_indices is not None else list(range(len(tables)))
            formatted_count = 0

            for idx in target_indices:
                if idx < 0 or idx >= len(tables):
                    continue
                tbl = tables[idx]
                formatted_count += 1
                tblPr = tbl.find("w:tblPr", namespaces={"w": W_NS})
                if tblPr is None:
                    tblPr = etree.Element(f"{{{W_NS}}}tblPr")
                    tbl.insert(0, tblPr)

                for old_s in tblPr.findall("w:tblStyle", namespaces={"w": W_NS}):
                    tblPr.remove(old_s)
                tblStyle = etree.Element(f"{{{W_NS}}}tblStyle")
                tblStyle.attrib[f"{{{W_NS}}}val"] = "TableGrid"
                tblPr.insert(0, tblStyle)

                for old_b in tblPr.findall("w:tblBorders", namespaces={"w": W_NS}):
                    tblPr.remove(old_b)
                borders = etree.SubElement(tblPr, f"{{{W_NS}}}tblBorders")
                for b_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
                    b_el = etree.SubElement(borders, f"{{{W_NS}}}{b_name}")
                    b_el.attrib[f"{{{W_NS}}}val"] = "single"
                    b_el.attrib[f"{{{W_NS}}}sz"] = str(border_sz)
                    b_el.attrib[f"{{{W_NS}}}space"] = "0"
                    b_el.attrib[f"{{{W_NS}}}color"] = str(border_color).replace("#", "")

                for old_m in tblPr.findall("w:tblCellMar", namespaces={"w": W_NS}):
                    tblPr.remove(old_m)
                cellMar = etree.SubElement(tblPr, f"{{{W_NS}}}tblCellMar")
                for m_name, m_val in [("top", "100"), ("bottom", "100"), ("left", "150"), ("right", "150")]:
                    m_el = etree.SubElement(cellMar, f"{{{W_NS}}}{m_name}")
                    m_el.attrib[f"{{{W_NS}}}w"] = m_val
                    m_el.attrib[f"{{{W_NS}}}type"] = "dxa"

                rows = tbl.findall("w:tr", namespaces={"w": W_NS})
                for r_idx, row in enumerate(rows):
                    trPr = row.find("w:trPr", namespaces={"w": W_NS})
                    if trPr is None:
                        trPr = etree.Element(f"{{{W_NS}}}trPr")
                        row.insert(0, trPr)

                    if cant_split and trPr.find("w:cantSplit", namespaces={"w": W_NS}) is None:
                        etree.SubElement(trPr, f"{{{W_NS}}}cantSplit")

                    if r_idx == 0 and repeat_header:
                        if trPr.find("w:tblHeader", namespaces={"w": W_NS}) is None:
                            etree.SubElement(trPr, f"{{{W_NS}}}tblHeader")

                    for cell in row.findall("w:tc", namespaces={"w": W_NS}):
                        tcPr = cell.find("w:tcPr", namespaces={"w": W_NS})
                        if tcPr is None:
                            tcPr = etree.Element(f"{{{W_NS}}}tcPr")
                            cell.insert(0, tcPr)

                        shd = tcPr.find("w:shd", namespaces={"w": W_NS})
                        if shd is None:
                            shd = etree.SubElement(tcPr, f"{{{W_NS}}}shd")
                        shd.attrib[f"{{{W_NS}}}val"] = "clear"
                        shd.attrib[f"{{{W_NS}}}color"] = "auto"
                        if r_idx == 0 and header_shading:
                            shd.attrib[f"{{{W_NS}}}fill"] = str(header_shading).replace("#", "")
                        else:
                            shd.attrib[f"{{{W_NS}}}fill"] = "FFFFFF"

                        if font_name or font_size_pt:
                            for cp in cell.findall(".//w:p", namespaces={"w": W_NS}):
                                for cr in cp.findall(".//w:r", namespaces={"w": W_NS}):
                                    crPr = cr.find("w:rPr", namespaces={"w": W_NS})
                                    if crPr is None:
                                        crPr = etree.SubElement(cr, f"{{{W_NS}}}rPr")
                                    if font_name:
                                        crf = crPr.find("w:rFonts", namespaces={"w": W_NS})
                                        if crf is None:
                                            crf = etree.SubElement(crPr, f"{{{W_NS}}}rFonts")
                                        crf.attrib[f"{{{W_NS}}}ascii"] = font_name
                                        crf.attrib[f"{{{W_NS}}}hAnsi"] = font_name
                                    if font_size_pt:
                                        sz_v = str(int(round(font_size_pt * 2)))
                                        for sz_tag in ["sz", "szCs"]:
                                            csz = crPr.find(f"w:{sz_tag}", namespaces={"w": W_NS})
                                            if csz is None:
                                                csz = etree.SubElement(crPr, f"{{{W_NS}}}{sz_tag}")
                                            csz.attrib[f"{{{W_NS}}}val"] = sz_v

            tmp_path = path.with_suffix(".tmp")
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in namelist:
                    if item == "word/document.xml":
                        zout.writestr(item, etree.tostring(doc_xml, xml_declaration=True, encoding="utf-8", standalone=True))
                    else:
                        zout.writestr(item, zin.read(item))

        if path.exists():
            os.remove(path)
        os.rename(tmp_path, path)

        return {
            "status": "success",
            "file_path": str(path),
            "tables_formatted": formatted_count,
            "total_tables": len(tables),
        }

    @staticmethod
    def clean_structure(
        file_path: str,
        remove_empty_paragraphs: bool = True,
        remove_empty_tables: bool = True,
        extract_single_image_tables: bool = True,
        max_consecutive_breaks: int = 1,
        normalize_all_colors_black: bool = False,
    ) -> Dict[str, Any]:
        """Automated hygiene primitive that strips trailing empty <w:p> nodes, removes empty tables with 0 text/drawings, and purges rogue page breaks."""
        path = normalize_path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        with open(path, "rb") as f:
            src_bytes = io.BytesIO(f.read())

        with zipfile.ZipFile(src_bytes, "r") as zin:
            namelist = zin.namelist()
            doc_xml = etree.fromstring(zin.read("word/document.xml"))
            body = doc_xml.find("w:body", namespaces={"w": W_NS})

            empty_tables_removed = 0
            images_extracted = 0
            empty_paragraphs_removed = 0
            page_breaks_cleaned = 0

            # 1. Table hygiene
            if remove_empty_tables or extract_single_image_tables:
                for tbl in list(body.findall("w:tbl", namespaces={"w": W_NS})):
                    text_nodes = tbl.findall(".//w:t", namespaces={"w": W_NS})
                    full_text = "".join(t.text for t in text_nodes if t.text).strip()
                    drawings = tbl.findall(".//w:drawing", namespaces={"w": W_NS})
                    picts = tbl.findall(".//w:pict", namespaces={"w": W_NS})
                    total_images = len(drawings) + len(picts)

                    if remove_empty_tables and not full_text and total_images == 0:
                        body.remove(tbl)
                        empty_tables_removed += 1
                        continue

                    if extract_single_image_tables and total_images > 0 and len(full_text) < 15:
                        tbl_parent = tbl.getparent()
                        tbl_pos = tbl_parent.index(tbl)
                        for d in drawings:
                            p_new = etree.Element(f"{{{W_NS}}}p")
                            pPr_new = etree.SubElement(p_new, f"{{{W_NS}}}pPr")
                            jc_new = etree.SubElement(pPr_new, f"{{{W_NS}}}jc")
                            jc_new.attrib[f"{{{W_NS}}}val"] = "center"
                            etree.SubElement(pPr_new, f"{{{W_NS}}}keepNext")
                            r_new = etree.SubElement(p_new, f"{{{W_NS}}}r")
                            d_parent = d.getparent()
                            if d_parent is not None:
                                d_parent.remove(d)
                            r_new.append(d)
                            tbl_parent.insert(tbl_pos, p_new)
                            tbl_pos += 1
                            images_extracted += 1

                        tbl_parent.remove(tbl)

            # 2. Page breaks hygiene
            if max_consecutive_breaks is not None:
                for p in body.findall(".//w:p", namespaces={"w": W_NS}):
                    brs = p.findall(".//w:br[@w:type='page']", namespaces={"w": W_NS})
                    if len(brs) > max_consecutive_breaks:
                        for br in brs[max_consecutive_breaks:]:
                            br.getparent().remove(br)
                            page_breaks_cleaned += 1

            # 3. Empty paragraphs cleanup
            if remove_empty_paragraphs:
                all_children = list(body)
                for child in all_children:
                    if child.tag.endswith("p"):
                        txt = "".join(t.text for t in child.findall(".//w:t", namespaces={"w": W_NS}) if t.text).strip()
                        drawings = child.findall(".//w:drawing", namespaces={"w": W_NS})
                        picts = child.findall(".//w:pict", namespaces={"w": W_NS})
                        brs = child.findall(".//w:br[@w:type='page']", namespaces={"w": W_NS})
                        sect_prs = child.findall(".//w:sectPr", namespaces={"w": W_NS})
                        if not txt and len(drawings) == 0 and len(picts) == 0 and len(brs) == 0 and len(sect_prs) == 0:
                            body.remove(child)
                            empty_paragraphs_removed += 1

            # 4. Color normalization
            if normalize_all_colors_black:
                for c in doc_xml.findall(".//w:color", namespaces={"w": W_NS}):
                    c.attrib[f"{{{W_NS}}}val"] = "000000"

            tmp_path = path.with_suffix(".tmp")
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in namelist:
                    if item == "word/document.xml":
                        zout.writestr(item, etree.tostring(doc_xml, xml_declaration=True, encoding="utf-8", standalone=True))
                    else:
                        zout.writestr(item, zin.read(item))

        if path.exists():
            os.remove(path)
        os.rename(tmp_path, path)

        return {
            "status": "success",
            "file_path": str(path),
            "empty_tables_removed": empty_tables_removed,
            "images_extracted_from_tables": images_extracted,
            "empty_paragraphs_removed": empty_paragraphs_removed,
            "page_breaks_cleaned": page_breaks_cleaned,
        }

