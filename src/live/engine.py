"""Live Microsoft Word backend engine via COM interface."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Callable, Dict, Generator, List, Literal, Optional, Tuple, Union

from src.core.constants import OUTLINE_LINE_LIMIT
from src.core.exceptions import ERROR_MESSAGES, LiveWordError
from src.core.paths import _norm, paths_match
from src.live.indexer import (
    WD_OUTLINE_LEVEL_BODY,
    WD_REVISION_DELETE,
    WD_WITHIN_TABLE,
    _build_visible_map,
    _com_idx,
    _get_local_outline,
    _visible_map,
)
from src.live.undo import live_call, tracking
from src.live.worker import HAS_COM, com_worker


def _hex_to_bgr(hex_str: Optional[str]) -> Optional[int]:
    """Convert hex color string (e.g. 'FF0000', '#FF0000') to Windows COLORREF integer (0xBBGGRR)."""
    if not hex_str:
        return None
    clean = hex_str.strip().lstrip("#")
    if len(clean) == 3:
        clean = "".join(c * 2 for c in clean)
    if len(clean) != 6:
        return None
    try:
        r = int(clean[0:2], 16)
        g = int(clean[2:4], 16)
        b = int(clean[4:6], 16)
        return r + (g << 8) + (b << 16)
    except ValueError:
        return None


class LiveWordEngine:
    """Live Word backend operations over Microsoft Word COM interface."""

    @staticmethod
    def read_document(
        doc: Any,
        file_path: Union[str, Path],
        mode: Literal["outline", "full", "table", "metadata"] = "full",
        table_index: Optional[int] = None,
    ) -> Union[str, Dict[str, Any]]:
        """Read and inspect open Word document structure and content."""
        def _read() -> Union[str, Dict[str, Any]]:
            path = Path(file_path)
            vmap = _build_visible_map(doc)
            body_indices = vmap["indices"]
            tables = doc.Tables
            num_tables = tables.Count
            num_paragraphs = len(body_indices)

            if mode == "metadata":
                props = {}
                try:
                    built_in = doc.BuiltInDocumentProperties
                    for key in ["Title", "Author", "Subject", "Keywords", "Last author"]:
                        try:
                            props[key] = str(built_in.Item(key).Value)
                        except Exception:
                            pass
                except Exception:
                    pass
                return {
                    "file_path": str(path),
                    "total_paragraphs": num_paragraphs,
                    "total_tables": num_tables,
                    "properties": props,
                }

            if mode == "table":
                if table_index is None:
                    raise ValueError("table_index (0-based) is required when mode='table'")
                if table_index < 0 or table_index >= num_tables:
                    raise IndexError(
                        f"Table index {table_index} is out of range. Document has {num_tables} tables (indices 0 to {num_tables - 1})."
                    )
                tbl = tables.Item(table_index + 1)
                matrix: List[List[str]] = []
                for r_idx in range(1, tbl.Rows.Count + 1):
                    row = tbl.Rows.Item(r_idx)
                    row_vals = [
                        row.Cells.Item(c).Range.Text.rstrip("\r\n\x07").replace("\n", " ").strip()
                        for c in range(1, row.Cells.Count + 1)
                    ]
                    matrix.append(row_vals)

                md_lines = [f"### Table `[table:{table_index}]` ({tbl.Rows.Count} rows x {tbl.Columns.Count} cols)", ""]
                if matrix:
                    header = matrix[0]
                    md_lines.append("| " + " | ".join(header) + " |")
                    md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")
                    for row in matrix[1:]:
                        md_lines.append("| " + " | ".join(row) + " |")
                return {
                    "table_index": table_index,
                    "rows_count": tbl.Rows.Count,
                    "columns_count": tbl.Columns.Count,
                    "matrix": matrix,
                    "markdown": "\n".join(md_lines),
                }

            if mode == "outline":
                lines = [
                    f"# Outline: {path.name}",
                    f"- **Total Paragraphs**: {num_paragraphs} (indices: [p:0] to [p:{max(0, num_paragraphs - 1)}])",
                    f"- **Total Tables**: {num_tables}",
                    "",
                    "### Structural Elements (0-based indices):",
                ]

                # Map table positions (1-based COM paragraph index where table starts)
                tbl_positions = []
                for t_idx in range(1, num_tables + 1):
                    try:
                        tr = tables.Item(t_idx).Range
                        p_start = doc.Range(0, tr.Start).Paragraphs.Count + 1
                        tbl_positions.append((p_start, t_idx))
                    except Exception:
                        tbl_positions.append((float("inf"), t_idx))

                # Combine body paragraphs and tables into a single document-order list
                elements = []
                for zero_idx, com_i in enumerate(body_indices):
                    elements.append((com_i, "p", zero_idx, com_i))
                for p_start, t_idx in tbl_positions:
                    elements.append((p_start, "tbl", t_idx - 1, t_idx))

                # Sort elements in document order
                elements.sort(key=lambda x: (x[0], 0 if x[1] == "tbl" else 1))

                # Fast heading detection: query doc.GetCrossReferenceItems(1)
                heading_com_indices = set()
                try:
                    items = doc.GetCrossReferenceItems(1)  # wdRefTypeHeading = 1
                except Exception:
                    items = ()

                if items:
                    rng = doc.Content
                    f = rng.Find
                    for heading_text in items:
                        clean = heading_text.strip()
                        if not clean:
                            continue
                        f.ClearFormatting()
                        if f.Execute(FindText=clean, Forward=True):
                            p_obj = rng.Paragraphs.Item(1)
                            c_i = doc.Range(0, p_obj.Range.End).Paragraphs.Count
                            heading_com_indices.add(c_i)
                            rng.Collapse(0)

                for elem in elements:
                    elem_type = elem[1]
                    if elem_type == "tbl":
                        tbl_idx = elem[2]
                        t_idx = elem[3]
                        try:
                            tbl = tables.Item(t_idx)
                            row_cnt = tbl.Rows.Count
                            col_cnt = tbl.Columns.Count
                            header_preview = []
                            if row_cnt > 0:
                                first_row = tbl.Rows.Item(1)
                                for c_i in range(1, min(col_cnt, 4) + 1):
                                    try:
                                        c_txt = first_row.Cells.Item(c_i).Range.Text.rstrip("\r\n\x07").strip()
                                        header_preview.append(c_txt)
                                    except Exception:
                                        pass
                            header_str = " | ".join(header_preview)
                            lines.append(f"- `[table:{tbl_idx}]` **Table ({row_cnt} rows x {col_cnt} cols)**: [{header_str}]")
                        except Exception:
                            lines.append(f"- `[table:{tbl_idx}]` **Table**")

                    elif elem_type == "p":
                        zero_idx = elem[2]
                        com_i = elem[3]

                        is_known_heading = com_i in heading_com_indices
                        if is_known_heading or len(lines) < OUTLINE_LINE_LIMIT or not items:
                            try:
                                p = doc.Paragraphs.Item(com_i)
                                style_name = str(p.Style.NameLocal) if p.Style else "Normal"
                                is_heading = is_known_heading or style_name.lower().startswith("heading") or style_name.lower() == "title"
                                text_snippet = p.Range.Text.rstrip("\r\n\x07").strip()
                                if is_heading or (len(lines) < OUTLINE_LINE_LIMIT and text_snippet):
                                    preview = (text_snippet[:80] + "...") if len(text_snippet) > 80 else text_snippet
                                    preview_text = f": \"{preview}\"" if preview else " (empty)"
                                    lines.append(f"- `[p:{zero_idx}]` **({style_name})**{preview_text}")
                            except Exception:
                                continue

                return "\n".join(lines)

            # Default: "full"
            lines = [f"# Document Content: {path.name}", ""]
            for zero_idx, com_i in enumerate(body_indices):
                p = doc.Paragraphs.Item(com_i)
                style_name = str(p.Style.NameLocal)
                txt = p.Range.Text.rstrip("\r\n\x07")
                lines.append(f"`[p:{zero_idx}]` ({style_name}): {txt}")

            for t_idx in range(1, num_tables + 1):
                tbl = tables.Item(t_idx)
                lines.append(f"\n`[table:{t_idx - 1}]` Table ({tbl.Rows.Count}x{tbl.Columns.Count}):")
                for r_idx in range(1, tbl.Rows.Count + 1):
                    row = tbl.Rows.Item(r_idx)
                    row_vals = [row.Cells.Item(c).Range.Text.rstrip("\r\n\x07") for c in range(1, row.Cells.Count + 1)]
                    lines.append(f"  row {r_idx - 1}: | " + " | ".join(row_vals) + " |")
                lines.append("")

            return "\n".join(lines)

        return com_worker.run(_read)

    @staticmethod
    def mutate_element(
        doc: Any,
        file_path: Union[str, Path],
        target_index: int,
        action: Literal["insert_before", "insert_after", "delete", "replace"],
        content: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Perform surgical mutation at target 0-based paragraph index with Track Changes."""
        def _mutate() -> Dict[str, Any]:
            vmap = _build_visible_map(doc)
            body_indices = vmap["indices"]
            if target_index < 0 or target_index >= len(body_indices):
                valid = f"0..{len(body_indices) - 1}" if body_indices else "empty"
                raise LiveWordError(
                    ERROR_MESSAGES["index_out_of_range"].format(target_index=target_index, valid_range=valid)
                )

            com_i = body_indices[target_index]
            p = doc.Paragraphs.Item(com_i)
            path = Path(file_path)

            with live_call(doc, "mutate_element", f"[p:{target_index}] {action}", track=True):
                if action == "delete":
                    p.Range.Delete()
                    fresh_outline = _get_local_outline(doc, target_index)
                    return {
                        "status": "success",
                        "action": "delete",
                        "target_index": target_index,
                        "file_path": str(path),
                        "local_outline": fresh_outline,
                    }

                if action == "replace":
                    c = content or {}
                    new_text = str(c.get("text", ""))
                    rng = p.Range
                    # Shrink range end by 1 to retain paragraph mark '\r'
                    if rng.End > rng.Start:
                        rng.End = rng.End - 1
                    rng.Text = new_text

                    if "style" in c:
                        try:
                            p.Style = c["style"]
                        except Exception:
                            pass
                    if c.get("bold") is not None:
                        rng.Font.Bold = bool(c["bold"])
                    if c.get("italic") is not None:
                        rng.Font.Italic = bool(c["italic"])

                    fresh_outline = _get_local_outline(doc, target_index)
                    return {
                        "status": "success",
                        "action": "replace",
                        "target_index": target_index,
                        "file_path": str(path),
                        "updated_text": new_text,
                        "local_outline": fresh_outline,
                    }

                if action in ("insert_before", "insert_after"):
                    c = content or {}
                    new_text = str(c.get("text", ""))
                    style = c.get("style", "Normal")

                    if action == "insert_before":
                        new_p = doc.Paragraphs.Add(p.Range)
                    else:
                        rng = p.Range
                        collapse_rng = doc.Range(rng.End, rng.End)
                        new_p = doc.Paragraphs.Add(collapse_rng)

                    try:
                        new_p.Style = style
                    except Exception:
                        pass

                    rng_new = new_p.Range
                    if rng_new.End > rng_new.Start:
                        rng_new.End = rng_new.End - 1
                    rng_new.Text = new_text

                    if c.get("bold") is not None:
                        rng_new.Font.Bold = bool(c["bold"])
                    if c.get("italic") is not None:
                        rng_new.Font.Italic = bool(c["italic"])

                    fresh_outline = _get_local_outline(doc, target_index)
                    return {
                        "status": "success",
                        "action": action,
                        "target_index": target_index,
                        "file_path": str(path),
                        "local_outline": fresh_outline,
                    }

            raise ValueError(f"Unknown action: {action}")

        return com_worker.run(_mutate)

    @staticmethod
    def find_and_replace(
        doc: Any,
        file_path: Union[str, Path],
        find_text: str,
        replace_text: str,
        match_case: bool = True,
        scope: Literal["all", "body", "tables", "headers_footers"] = "body",
    ) -> Dict[str, Any]:
        """Locate-then-set find and replace avoiding 255-char replacement limits and skipping deletions."""
        if not find_text:
            raise ValueError("find_text cannot be empty.")
        if len(find_text) > 255:
            raise ValueError("find_text exceeds Word 255 character limit.")

        def _fnr() -> Dict[str, Any]:
            path = Path(file_path)
            count = 0
            with live_call(doc, "find_and_replace", f"'{find_text[:20]}'", track=True):
                ranges = []
                if scope in ("all", "body"):
                    ranges.append(doc.Content)
                elif scope == "tables":
                    for t_i in range(1, doc.Tables.Count + 1):
                        ranges.append(doc.Tables.Item(t_i).Range)

                if scope in ("all", "headers_footers"):
                    for s_i in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections.Item(s_i)
                        for hf_group in (sec.Headers, sec.Footers):
                            for hf_i in range(1, hf_group.Count + 1):
                                try:
                                    hf = hf_group.Item(hf_i)
                                    if getattr(hf, "Exists", False):
                                        ranges.append(hf.Range)
                                except Exception:
                                    pass

                for curr_rng in ranges:
                    find_obj = curr_rng.Find
                    find_obj.ClearFormatting()
                    find_obj.Text = find_text
                    find_obj.MatchCase = match_case
                    find_obj.Forward = True
                    find_obj.Wrap = 0  # wdFindStop

                    while find_obj.Execute():
                        in_del = False
                        try:
                            for rev in curr_rng.Revisions:
                                if rev.Type == WD_REVISION_DELETE:
                                    in_del = True
                                    break
                        except Exception:
                            pass

                        if not in_del:
                            curr_rng.Text = replace_text
                            count += 1

                        curr_rng.Collapse(0)  # wdCollapseEnd

            return {
                "status": "success",
                "file_path": str(path),
                "matches_replaced": count,
                "find_text": find_text,
                "replace_text": replace_text,
            }

        return com_worker.run(_fnr)

    @staticmethod
    def set_page_layout(
        doc: Any,
        file_path: Union[str, Path],
        top_cm: float = 2.54,
        bottom_cm: float = 2.54,
        left_cm: float = 2.54,
        right_cm: float = 2.54,
        page_size: Literal["letter", "a4", "legal"] = "letter",
        orientation: Literal["portrait", "landscape"] = "portrait",
    ) -> Dict[str, Any]:
        """Modify margins and dimensions across all sections in the active Word document."""
        def _set() -> Dict[str, Any]:
            path = Path(file_path)
            cm_to_pt = 72.0 / 2.54
            top_pt = top_cm * cm_to_pt
            bottom_pt = bottom_cm * cm_to_pt
            left_pt = left_cm * cm_to_pt
            right_pt = right_cm * cm_to_pt

            sizes = {
                "letter": (612.0, 792.0),
                "a4": (595.3, 841.9),
                "legal": (612.0, 1008.0),
            }
            w_pt, h_pt = sizes.get(page_size.lower(), sizes["letter"])
            if orientation.lower() == "landscape":
                w_pt, h_pt = h_pt, w_pt
                orient_val = 1  # wdOrientLandscape
            else:
                orient_val = 0  # wdOrientPortrait

            with live_call(doc, "set_page_layout", f"{page_size} {orientation}"):
                sections = doc.Sections
                count = sections.Count
                for s_idx in range(1, count + 1):
                    ps = sections.Item(s_idx).PageSetup
                    ps.Orientation = orient_val
                    ps.PageWidth = w_pt
                    ps.PageHeight = h_pt
                    ps.TopMargin = top_pt
                    ps.BottomMargin = bottom_pt
                    ps.LeftMargin = left_pt
                    ps.RightMargin = right_pt

            return {
                "status": "success",
                "file_path": str(path),
                "sections_updated": count,
                "margins_cm": {"top": top_cm, "bottom": bottom_cm, "left": left_cm, "right": right_cm},
                "page_size": page_size,
                "orientation": orientation,
            }

        return com_worker.run(_set)

    @staticmethod
    def update_styles(
        doc: Any,
        file_path: Union[str, Path],
        styles_config: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Any]:
        """Configure font, size, bold, italic, line spacing, space before/after, and color on document styles."""
        def _update() -> Dict[str, Any]:
            path = Path(file_path)
            updated_styles = []

            align_map = {
                "left": 0,      # wdAlignParagraphLeft
                "center": 1,    # wdAlignParagraphCenter
                "right": 2,     # wdAlignParagraphRight
                "justify": 3,   # wdAlignParagraphJustify
            }

            with live_call(doc, "update_styles", f"{len(styles_config)} styles"):
                for style_name, cfg in styles_config.items():
                    try:
                        style_obj = doc.Styles(style_name)
                    except Exception:
                        try:
                            style_obj = doc.Styles.Add(Name=style_name, Type=1)
                        except Exception:
                            continue

                    font = style_obj.Font
                    pf = style_obj.ParagraphFormat

                    if "font" in cfg and cfg["font"]:
                        font.Name = str(cfg["font"])
                    if "size_pt" in cfg and cfg["size_pt"] is not None:
                        font.Size = float(cfg["size_pt"])
                    if "bold" in cfg and cfg["bold"] is not None:
                        font.Bold = bool(cfg["bold"])
                    if "italic" in cfg and cfg["italic"] is not None:
                        font.Italic = bool(cfg["italic"])
                    if "color" in cfg and cfg["color"] is not None:
                        c_val = _hex_to_bgr(str(cfg["color"]))
                        if c_val is not None:
                            font.Color = c_val

                    if "alignment" in cfg and cfg["alignment"] in align_map:
                        pf.Alignment = align_map[cfg["alignment"]]
                    if "space_after_pt" in cfg and cfg["space_after_pt"] is not None:
                        pf.SpaceAfter = float(cfg["space_after_pt"])
                    if "space_before_pt" in cfg and cfg["space_before_pt"] is not None:
                        pf.SpaceBefore = float(cfg["space_before_pt"])
                    if "line_spacing" in cfg and cfg["line_spacing"] is not None:
                        ls = float(cfg["line_spacing"])
                        if abs(ls - 1.0) < 0.05:
                            pf.LineSpacingRule = 0
                        elif abs(ls - 1.5) < 0.05:
                            pf.LineSpacingRule = 1
                        elif abs(ls - 2.0) < 0.05:
                            pf.LineSpacingRule = 2
                        else:
                            pf.LineSpacingRule = 5
                            pf.LineSpacing = ls * 12.0
                    if "keep_next" in cfg and cfg["keep_next"] is not None:
                        pf.KeepWithNext = bool(cfg["keep_next"])

                    updated_styles.append(style_name)

            return {
                "status": "success",
                "file_path": str(path),
                "styles_updated_count": len(updated_styles),
                "updated_styles": updated_styles,
            }

        return com_worker.run(_update)

    @staticmethod
    def compose_document(
        doc: Any,
        file_path: Union[str, Path],
        actions: List[Dict[str, Any]],
        create_if_missing: bool = True,
    ) -> Dict[str, Any]:
        """Append a batch of formatted elements to the active Word document."""
        def _compose() -> Dict[str, Any]:
            path = Path(file_path)
            executed: List[str] = []

            with live_call(doc, "compose_document", f"{len(actions)} actions"):
                for idx, act in enumerate(actions):
                    act_type = str(act.get("type", "")).lower()

                    if act_type == "heading":
                        text = str(act.get("text", ""))
                        level = int(act.get("level", 1))
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        p = doc.Paragraphs.Add(Range=end_rng)
                        p.Range.Text = text + "\n"
                        try:
                            p.Style = f"Heading {level}"
                        except Exception:
                            pass
                        executed.append(f"#{idx} Added Heading (level {level}): \"{text[:40]}\"")

                    elif act_type in ("paragraph", "list_item"):
                        text = str(act.get("text", ""))
                        default_style = "List Bullet" if act_type == "list_item" else "Normal"
                        style = str(act.get("style", default_style))
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        p = doc.Paragraphs.Add(Range=end_rng)
                        p.Range.Text = text + "\n"
                        try:
                            p.Style = style
                        except Exception:
                            pass
                        if act.get("bold") is not None:
                            p.Range.Font.Bold = bool(act["bold"])
                        if act.get("italic") is not None:
                            p.Range.Font.Italic = bool(act["italic"])
                        executed.append(f"#{idx} Added Paragraph: \"{text[:40]}\"")

                    elif act_type == "table":
                        headers = act.get("headers", [])
                        rows = act.get("rows", [])
                        num_rows = len(rows) + (1 if headers else 0)
                        num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        tbl = doc.Tables.Add(Range=end_rng, NumRows=num_rows, NumColumns=num_cols)
                        try:
                            tbl.Style = act.get("style", "Table Grid")
                        except Exception:
                            pass
                        curr_r = 1
                        if headers:
                            for c_i, h in enumerate(headers):
                                tbl.Cell(1, c_i + 1).Range.Text = str(h)
                            curr_r = 2
                        for r_vals in rows:
                            for c_i, val in enumerate(r_vals):
                                if c_i < num_cols:
                                    tbl.Cell(curr_r, c_i + 1).Range.Text = str(val)
                            curr_r += 1
                        executed.append(f"#{idx} Added Table ({num_rows}x{num_cols})")

                    elif act_type == "page_break":
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        end_rng.InsertBreak(7)  # wdPageBreak = 7
                        executed.append(f"#{idx} Added Page Break")

            return {
                "status": "success",
                "file_path": str(path),
                "total_actions": len(actions),
                "actions_executed": executed,
                "paragraphs_count": doc.Paragraphs.Count,
                "tables_count": doc.Tables.Count,
            }

        return com_worker.run(_compose)

    @staticmethod
    def format_table_grid(
        doc: Any,
        file_path: Union[str, Path],
        table_indices: Optional[List[int]] = None,
        border_sz: int = 4,
        border_color: str = "000000",
        header_shading: Optional[str] = "F2F2F2",
        cant_split: bool = True,
        repeat_header: bool = True,
        font_name: Optional[str] = "Times New Roman",
        font_size_pt: Optional[float] = 11.0,
    ) -> Dict[str, Any]:
        """Normalize borders, header shading, pagination, and font for tables in the active Word document."""
        def _format() -> Dict[str, Any]:
            path = Path(file_path)
            tables = doc.Tables
            total_tables = tables.Count
            target_indices = table_indices if table_indices is not None else list(range(total_tables))
            formatted_count = 0

            bgr_color = _hex_to_bgr(border_color) or 0
            bgr_shading = _hex_to_bgr(header_shading) if header_shading else None

            with live_call(doc, "format_table_grid", f"{len(target_indices)} tables"):
                for idx in target_indices:
                    if idx < 0 or idx >= total_tables:
                        continue
                    tbl = tables.Item(idx + 1)
                    formatted_count += 1

                    try:
                        tbl.Borders.Enable = True
                        tbl.Borders.OutsideLineStyle = 1  # wdLineStyleSingle
                        tbl.Borders.InsideLineStyle = 1
                        tbl.Borders.OutsideLineWidth = border_sz
                        tbl.Borders.InsideLineWidth = border_sz
                        tbl.Borders.OutsideColor = bgr_color
                        tbl.Borders.InsideColor = bgr_color
                    except Exception:
                        pass

                    try:
                        tbl.TopPadding = 5.0
                        tbl.BottomPadding = 5.0
                        tbl.LeftPadding = 7.5
                        tbl.RightPadding = 7.5
                    except Exception:
                        pass

                    row_count = tbl.Rows.Count
                    for r_idx in range(1, row_count + 1):
                        row = tbl.Rows.Item(r_idx)
                        if cant_split:
                            try:
                                row.AllowBreakAcrossPages = 0  # False
                            except Exception:
                                pass

                        if r_idx == 1 and repeat_header:
                            try:
                                row.HeadingFormat = -1  # True
                            except Exception:
                                pass

                        if r_idx == 1 and bgr_shading is not None:
                            try:
                                row.Shading.BackgroundPatternColor = bgr_shading
                            except Exception:
                                pass
                        elif bgr_shading is not None:
                            try:
                                row.Shading.BackgroundPatternColor = 0xFFFFFF
                            except Exception:
                                pass

                    if font_name or font_size_pt:
                        try:
                            rng = tbl.Range
                            if font_name:
                                rng.Font.Name = font_name
                            if font_size_pt:
                                rng.Font.Size = font_size_pt
                        except Exception:
                            pass

            return {
                "status": "success",
                "file_path": str(path),
                "tables_formatted": formatted_count,
                "total_tables": total_tables,
            }

        return com_worker.run(_format)

    @staticmethod
    def clean_structure(
        doc: Any,
        file_path: Union[str, Path],
        remove_empty_paragraphs: bool = True,
        remove_empty_tables: bool = True,
        extract_single_image_tables: bool = True,
        max_consecutive_breaks: int = 1,
        normalize_all_colors_black: bool = False,
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """Automated hygiene primitive for active Word document removing empty nodes and rogue breaks."""
        def _clean() -> Dict[str, Any]:
            path = Path(file_path)
            empty_tables_removed = 0
            images_extracted = 0
            empty_paragraphs_removed = 0
            page_breaks_cleaned = 0

            with live_call(doc, "clean_structure", f"dry_run={dry_run}"):
                if remove_empty_tables:
                    tbl_count = doc.Tables.Count
                    for t_idx in range(tbl_count, 0, -1):
                        try:
                            tbl = doc.Tables.Item(t_idx)
                            t_text = tbl.Range.Text.rstrip("\r\n\x07").replace("\x07", "").strip()
                            inline_shapes = tbl.Range.InlineShapes.Count
                            if not t_text and inline_shapes == 0:
                                if not dry_run:
                                    tbl.Delete()
                                empty_tables_removed += 1
                        except Exception:
                            pass

                if remove_empty_paragraphs:
                    p_count = doc.Paragraphs.Count
                    for p_i in range(p_count, 0, -1):
                        try:
                            p = doc.Paragraphs.Item(p_i)
                            if p.Range.Information(WD_WITHIN_TABLE):
                                continue
                            if p_i == 1 and p.Range.End >= doc.Content.End:
                                continue
                            txt = p.Range.Text.rstrip("\r\n\x07").strip()
                            if not txt and p.Range.InlineShapes.Count == 0 and p.Range.ShapeRange.Count == 0:
                                if not dry_run:
                                    p.Range.Delete()
                                empty_paragraphs_removed += 1
                        except Exception:
                            pass

                if normalize_all_colors_black and not dry_run:
                    try:
                        doc.Content.Font.Color = 0  # wdColorBlack
                    except Exception:
                        pass

            return {
                "status": "success",
                "file_path": str(path),
                "empty_tables_removed": empty_tables_removed,
                "images_extracted_from_tables": images_extracted,
                "empty_paragraphs_removed": empty_paragraphs_removed,
                "page_breaks_cleaned": page_breaks_cleaned,
                "dry_run": dry_run,
            }

        return com_worker.run(_clean)
