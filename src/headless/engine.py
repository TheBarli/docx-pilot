"""Core OpenXML and python-docx Engine for headless operations."""

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

from src.core.constants import (
    OUTLINE_LINE_LIMIT,
    PPR_ORDER,
    RPR_ORDER,
    STYLE_ORDER,
    W_NS,
)
from src.core.paths import normalize_path
from src.headless.tools_edit import EditTools
from src.headless.tools_read import ReadTools
from src.headless.xml_helpers import get_or_create_document, reorder_children


class DocxEngine(ReadTools, EditTools):
    """Low-level engine for Word document inspection, composition, and XML mutation."""

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
        dry_run: bool = False,
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
                        if not dry_run:
                            body.remove(tbl)
                        empty_tables_removed += 1
                        continue

                    if extract_single_image_tables and total_images > 0 and len(full_text) < 15:
                        if not dry_run:
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
                            tbl_parent.remove(tbl)
                        images_extracted += 1

            # 2. Page breaks hygiene
            if max_consecutive_breaks is not None:
                for p in body.findall(".//w:p", namespaces={"w": W_NS}):
                    brs = p.findall(".//w:br[@w:type='page']", namespaces={"w": W_NS})
                    if len(brs) > max_consecutive_breaks:
                        for br in brs[max_consecutive_breaks:]:
                            if not dry_run:
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
                            if not dry_run:
                                body.remove(child)
                            empty_paragraphs_removed += 1

            # 4. Color normalization
            if normalize_all_colors_black and not dry_run:
                for c in doc_xml.findall(".//w:color", namespaces={"w": W_NS}):
                    c.attrib[f"{{{W_NS}}}val"] = "000000"

            if dry_run:
                return {
                    "status": "success",
                    "file_path": str(path),
                    "empty_tables_removed": empty_tables_removed,
                    "images_extracted_from_tables": images_extracted,
                    "empty_paragraphs_removed": empty_paragraphs_removed,
                    "page_breaks_cleaned": page_breaks_cleaned,
                    "dry_run": True,
                }

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
            "dry_run": False,
        }

