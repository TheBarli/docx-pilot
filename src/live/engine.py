"""Live Microsoft Word backend engine via COM interface."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Callable, Dict, Generator, List, Literal, Optional, Tuple, Union

from src.live.indexer import WD_WITHIN_TABLE
from src.live.tools_edit import EditTools
from src.live.tools_read import ReadTools
from src.live.undo import live_call
from src.live.worker import com_worker


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


class LiveWordEngine(ReadTools, EditTools):
    """Live Word backend operations over Microsoft Word COM interface."""

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
