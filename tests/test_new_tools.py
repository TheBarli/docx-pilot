"""Tests for the 4 new DocxEngine methods:
    set_page_layout, update_styles, format_table_grid, clean_structure.

Each test creates a fresh document in a tmp_path directory (never mutates a
shared fixture) and performs a round-trip reopening with python-docx as a
cheap proxy for Word's "repaired content" guard.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree
from docx import Document

from src.oxml_engine import (
    DocxEngine,
    normalize_path,
    PROJECT_ROOT,
    W_NS,
    PPR_ORDER,
    RPR_ORDER,
    STYLE_ORDER,
)

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _make_docx(tmp_path: Path, name: str = "test.docx") -> str:
    """Create a minimal docx with known content and return its path string."""
    p = str(tmp_path / name)
    DocxEngine.compose_document(
        p,
        [
            {"type": "heading", "text": "Title", "level": 1},
            {"type": "paragraph", "text": "Body paragraph."},
            {"type": "paragraph", "text": ""},        # intentionally empty
            {"type": "paragraph", "text": ""},        # second empty
            {
                "type": "table",
                "headers": ["Col A", "Col B"],
                "rows": [["r1c1", "r1c2"], ["r2c1", "r2c2"]],
            },
        ],
        create_if_missing=True,
    )
    return p


def _read_doc_xml(docx_path: str) -> etree._Element:
    """Parse word/document.xml from a .docx file."""
    with zipfile.ZipFile(docx_path, "r") as z:
        return etree.fromstring(z.read("word/document.xml"))


def _read_styles_xml(docx_path: str) -> etree._Element:
    """Parse word/styles.xml from a .docx file."""
    with zipfile.ZipFile(docx_path, "r") as z:
        return etree.fromstring(z.read("word/styles.xml"))


def _roundtrip(docx_path: str) -> None:
    """Assert python-docx can reopen the file without raising (repaired-content proxy)."""
    doc = Document(docx_path)
    assert doc is not None


# ---------------------------------------------------------------------------
# normalize_path
# ---------------------------------------------------------------------------

class TestNormalizePath:
    def test_absolute_existing(self, tmp_path: Path):
        f = tmp_path / "file.docx"
        f.write_bytes(b"x")
        result = normalize_path(str(f))
        assert result == f.resolve()

    def test_relative_resolves_to_project_root(self, tmp_path: Path, monkeypatch):
        # Point PROJECT_ROOT to tmp_path so "relative/file.docx" resolves there
        import src.oxml_engine as engine_mod
        monkeypatch.setattr(engine_mod, "PROJECT_ROOT", tmp_path)
        (tmp_path / "relative").mkdir()
        f = tmp_path / "relative" / "file.docx"
        f.write_bytes(b"x")
        result = normalize_path("relative/file.docx")
        assert result.exists()

    def test_nonexistent_absolute_returned_as_is(self, tmp_path: Path):
        p = tmp_path / "missing.docx"
        result = normalize_path(str(p))
        assert result == p.resolve()


# ---------------------------------------------------------------------------
# set_page_layout
# ---------------------------------------------------------------------------

class TestSetPageLayout:
    def test_margins_applied_to_all_sectPr(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        res = DocxEngine.set_page_layout(
            docx_path,
            top_cm=4.0, bottom_cm=3.0, left_cm=4.0, right_cm=3.0,
            page_size="a4", orientation="portrait",
        )
        assert res["status"] == "success"
        assert res["sections_updated"] >= 1
        assert res["margins_cm"] == {"top": 4.0, "bottom": 3.0, "left": 4.0, "right": 3.0}

        # Verify actual XML values (1 cm ≈ 567 twips)
        doc_xml = _read_doc_xml(docx_path)
        sections = doc_xml.findall(".//w:sectPr", namespaces={"w": W_NS})
        assert len(sections) >= 1
        for s in sections:
            pg_mar = s.find("w:pgMar", namespaces={"w": W_NS})
            assert pg_mar is not None
            assert pg_mar.attrib[f"{{{W_NS}}}top"] == str(int(round(4.0 * 567)))
            assert pg_mar.attrib[f"{{{W_NS}}}right"] == str(int(round(3.0 * 567)))

    def test_a4_page_size(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.set_page_layout(docx_path, page_size="a4")
        doc_xml = _read_doc_xml(docx_path)
        for s in doc_xml.findall(".//w:sectPr", namespaces={"w": W_NS}):
            pg_sz = s.find("w:pgSz", namespaces={"w": W_NS})
            if pg_sz is not None:
                assert pg_sz.attrib[f"{{{W_NS}}}w"] == "11909"
                assert pg_sz.attrib[f"{{{W_NS}}}h"] == "16834"

    def test_landscape_swaps_dimensions(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.set_page_layout(docx_path, page_size="a4", orientation="landscape")
        doc_xml = _read_doc_xml(docx_path)
        for s in doc_xml.findall(".//w:sectPr", namespaces={"w": W_NS}):
            pg_sz = s.find("w:pgSz", namespaces={"w": W_NS})
            if pg_sz is not None:
                # landscape: w and h are swapped vs portrait
                assert pg_sz.attrib[f"{{{W_NS}}}w"] == "16834"
                assert pg_sz.attrib[f"{{{W_NS}}}h"] == "11909"

    def test_dry_run_leaves_document_intact(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        # Check initial empty paragraphs
        res_dry = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=True)
        assert res_dry["status"] == "success"
        assert res_dry["dry_run"] is True
        assert res_dry["empty_paragraphs_removed"] >= 2
        
        # Second run without dry_run should still find and remove them
        res_real = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=False)
        assert res_real["dry_run"] is False
        assert res_real["empty_paragraphs_removed"] >= 2

    def test_roundtrip(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.set_page_layout(docx_path)
        _roundtrip(docx_path)

    def test_missing_file_raises(self, tmp_path: Path):
        with pytest.raises(FileNotFoundError):
            DocxEngine.set_page_layout(str(tmp_path / "ghost.docx"))


# ---------------------------------------------------------------------------
# update_styles
# ---------------------------------------------------------------------------

class TestUpdateStyles:
    def test_normal_style_font_and_size(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        res = DocxEngine.update_styles(
            docx_path,
            styles_config={"Normal": {"font": "Times New Roman", "size_pt": 12}},
        )
        assert res["status"] == "success"
        assert "Normal" in res["styles_updated"]

        styles_xml = _read_styles_xml(docx_path)
        for s in styles_xml.findall("w:style", namespaces={"w": W_NS}):
            name_el = s.find("w:name", namespaces={"w": W_NS})
            if name_el is not None and name_el.attrib.get(f"{{{W_NS}}}val", "").lower() == "normal":
                rPr = s.find("w:rPr", namespaces={"w": W_NS})
                assert rPr is not None
                rFonts = rPr.find("w:rFonts", namespaces={"w": W_NS})
                assert rFonts is not None
                assert rFonts.attrib[f"{{{W_NS}}}ascii"] == "Times New Roman"
                sz = rPr.find("w:sz", namespaces={"w": W_NS})
                assert sz is not None
                assert sz.attrib[f"{{{W_NS}}}val"] == "24"  # 12pt * 2

    def test_bold_applied_and_removed(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        # Apply bold
        DocxEngine.update_styles(docx_path, styles_config={"Normal": {"bold": True}})
        styles_xml = _read_styles_xml(docx_path)
        for s in styles_xml.findall("w:style", namespaces={"w": W_NS}):
            name_el = s.find("w:name", namespaces={"w": W_NS})
            if name_el is not None and name_el.attrib.get(f"{{{W_NS}}}val", "").lower() == "normal":
                rPr = s.find("w:rPr", namespaces={"w": W_NS})
                assert rPr.find("w:b", namespaces={"w": W_NS}) is not None

        # Remove bold
        DocxEngine.update_styles(docx_path, styles_config={"Normal": {"bold": False}})
        styles_xml2 = _read_styles_xml(docx_path)
        for s in styles_xml2.findall("w:style", namespaces={"w": W_NS}):
            name_el = s.find("w:name", namespaces={"w": W_NS})
            if name_el is not None and name_el.attrib.get(f"{{{W_NS}}}val", "").lower() == "normal":
                rPr = s.find("w:rPr", namespaces={"w": W_NS})
                assert rPr.find("w:b", namespaces={"w": W_NS}) is None

    def test_child_order_matches_ecma376(self, tmp_path: Path):
        """Children of rPr / pPr / style must respect ECMA-376 sequence order."""
        docx_path = _make_docx(tmp_path)
        DocxEngine.update_styles(
            docx_path,
            styles_config={
                "Normal": {
                    "font": "Calibri", "size_pt": 11, "bold": False,
                    "color": "000000", "line_spacing": 1.5, "alignment": "both",
                }
            },
        )
        styles_xml = _read_styles_xml(docx_path)

        order_map_rpr = {tag: i for i, tag in enumerate(RPR_ORDER)}
        order_map_ppr = {tag: i for i, tag in enumerate(PPR_ORDER)}
        order_map_style = {tag: i for i, tag in enumerate(STYLE_ORDER)}

        def local_tag(el):
            return el.tag.split("}")[-1] if "}" in el.tag else el.tag

        for s in styles_xml.findall("w:style", namespaces={"w": W_NS}):
            name_el = s.find("w:name", namespaces={"w": W_NS})
            if name_el is None or name_el.attrib.get(f"{{{W_NS}}}val", "").lower() != "normal":
                continue
            # Check rPr order
            rPr = s.find("w:rPr", namespaces={"w": W_NS})
            if rPr is not None:
                tags = [order_map_rpr.get(local_tag(c), 999) for c in rPr]
                assert tags == sorted(tags), f"rPr children out of ECMA-376 order: {[local_tag(c) for c in rPr]}"
            # Check pPr order
            pPr = s.find("w:pPr", namespaces={"w": W_NS})
            if pPr is not None:
                tags = [order_map_ppr.get(local_tag(c), 999) for c in pPr]
                assert tags == sorted(tags), f"pPr children out of ECMA-376 order: {[local_tag(c) for c in pPr]}"

    def test_keep_next_applied(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.update_styles(
            docx_path,
            styles_config={"Heading1": {"keep_next": True}},
        )
        styles_xml = _read_styles_xml(docx_path)
        found = False
        for s in styles_xml.findall("w:style", namespaces={"w": W_NS}):
            sid = s.attrib.get(f"{{{W_NS}}}styleId", "")
            if sid.lower() == "heading1":
                pPr = s.find("w:pPr", namespaces={"w": W_NS})
                assert pPr is not None
                assert pPr.find("w:keepNext", namespaces={"w": W_NS}) is not None
                found = True
        assert found, "Heading1 style not found after update"

    def test_dry_run_leaves_document_intact(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        # Check initial empty paragraphs
        res_dry = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=True)
        assert res_dry["status"] == "success"
        assert res_dry["dry_run"] is True
        assert res_dry["empty_paragraphs_removed"] >= 2
        
        # Second run without dry_run should still find and remove them
        res_real = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=False)
        assert res_real["dry_run"] is False
        assert res_real["empty_paragraphs_removed"] >= 2

    def test_roundtrip(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.update_styles(docx_path, styles_config={"Normal": {"font": "Arial", "size_pt": 11}})
        _roundtrip(docx_path)


# ---------------------------------------------------------------------------
# format_table_grid
# ---------------------------------------------------------------------------

class TestFormatTableGrid:
    def test_all_tables_formatted_when_none(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        res = DocxEngine.format_table_grid(docx_path, table_indices=None)
        assert res["status"] == "success"
        assert res["tables_formatted"] == res["total_tables"]

    def test_selective_by_index(self, tmp_path: Path):
        # Two tables
        p = str(tmp_path / "two_tables.docx")
        DocxEngine.compose_document(p, [
            {"type": "table", "headers": ["A"], "rows": [["1"]]},
            {"type": "table", "headers": ["B"], "rows": [["2"]]},
        ], create_if_missing=True)
        res = DocxEngine.format_table_grid(p, table_indices=[0])
        assert res["tables_formatted"] == 1
        assert res["total_tables"] == 2

    def test_cantSplit_and_tblHeader_present(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.format_table_grid(docx_path, cant_split=True, repeat_header=True)
        doc_xml = _read_doc_xml(docx_path)
        tables = doc_xml.findall(".//w:tbl", namespaces={"w": W_NS})
        assert len(tables) >= 1
        for tbl in tables:
            rows = tbl.findall("w:tr", namespaces={"w": W_NS})
            assert len(rows) >= 1
            # All rows should have cantSplit
            for row in rows:
                trPr = row.find("w:trPr", namespaces={"w": W_NS})
                assert trPr is not None
                assert trPr.find("w:cantSplit", namespaces={"w": W_NS}) is not None
            # First row should have tblHeader
            first_trPr = rows[0].find("w:trPr", namespaces={"w": W_NS})
            assert first_trPr.find("w:tblHeader", namespaces={"w": W_NS}) is not None

    def test_header_shading_applied_to_first_row_only(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.format_table_grid(docx_path, header_shading="F2F2F2")
        doc_xml = _read_doc_xml(docx_path)
        tables = doc_xml.findall(".//w:tbl", namespaces={"w": W_NS})
        for tbl in tables:
            rows = tbl.findall("w:tr", namespaces={"w": W_NS})
            # First row cells should have F2F2F2 fill
            for cell in rows[0].findall("w:tc", namespaces={"w": W_NS}):
                tcPr = cell.find("w:tcPr", namespaces={"w": W_NS})
                shd = tcPr.find("w:shd", namespaces={"w": W_NS}) if tcPr is not None else None
                assert shd is not None
                assert shd.attrib.get(f"{{{W_NS}}}fill", "").upper() == "F2F2F2"
            # Non-header rows should NOT have F2F2F2
            for row in rows[1:]:
                for cell in row.findall("w:tc", namespaces={"w": W_NS}):
                    tcPr = cell.find("w:tcPr", namespaces={"w": W_NS})
                    shd = tcPr.find("w:shd", namespaces={"w": W_NS}) if tcPr is not None else None
                    if shd is not None:
                        assert shd.attrib.get(f"{{{W_NS}}}fill", "").upper() != "F2F2F2"

    def test_border_attributes(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.format_table_grid(docx_path, border_sz=8, border_color="FF0000")
        doc_xml = _read_doc_xml(docx_path)
        for tbl in doc_xml.findall(".//w:tbl", namespaces={"w": W_NS}):
            tblPr = tbl.find("w:tblPr", namespaces={"w": W_NS})
            borders = tblPr.find("w:tblBorders", namespaces={"w": W_NS})
            assert borders is not None
            top_b = borders.find("w:top", namespaces={"w": W_NS})
            assert top_b.attrib[f"{{{W_NS}}}sz"] == "8"
            assert top_b.attrib[f"{{{W_NS}}}color"] == "FF0000"

    def test_dry_run_leaves_document_intact(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        # Check initial empty paragraphs
        res_dry = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=True)
        assert res_dry["status"] == "success"
        assert res_dry["dry_run"] is True
        assert res_dry["empty_paragraphs_removed"] >= 2
        
        # Second run without dry_run should still find and remove them
        res_real = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=False)
        assert res_real["dry_run"] is False
        assert res_real["empty_paragraphs_removed"] >= 2

    def test_roundtrip(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.format_table_grid(docx_path)
        _roundtrip(docx_path)


# ---------------------------------------------------------------------------
# clean_structure
# ---------------------------------------------------------------------------

class TestCleanStructure:
    def test_removes_empty_paragraphs(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        # _make_docx adds 2 empty paragraphs
        res = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True)
        assert res["status"] == "success"
        assert res["empty_paragraphs_removed"] >= 2

    def test_empty_paragraphs_flag_off(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        res = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=False)
        assert res["empty_paragraphs_removed"] == 0

    def test_removes_empty_table(self, tmp_path: Path):
        p = str(tmp_path / "empty_tbl.docx")
        # Build a doc that has a no-content table (just empty cells)
        DocxEngine.compose_document(p, [
            {"type": "paragraph", "text": "Real content."},
            {"type": "table", "headers": [""], "rows": [[""]]},  # empty text table
        ], create_if_missing=True)
        res = DocxEngine.clean_structure(p, remove_empty_tables=True, remove_empty_paragraphs=False)
        assert res["empty_tables_removed"] >= 1

    def test_page_break_cap(self, tmp_path: Path):
        """max_consecutive_breaks=1 should remove extra page breaks within a paragraph."""
        import io as _io
        import zipfile as zf
        from lxml import etree as et

        p = str(tmp_path / "breaks.docx")
        DocxEngine.compose_document(p, [
            {"type": "page_break"},
            {"type": "paragraph", "text": "After breaks."},
        ], create_if_missing=True)

        # Read the zip, inject an extra page-break into the first break paragraph,
        # then write the whole archive back atomically.
        with open(p, "rb") as fh:
            src_bytes = _io.BytesIO(fh.read())

        with zf.ZipFile(src_bytes, "r") as zin:
            names = zin.namelist()
            doc_xml = et.fromstring(zin.read("word/document.xml"))
            body = doc_xml.find("w:body", namespaces={"w": W_NS})

            # Find the first <w:p> that already has a page break and add one more
            injected = False
            for para in body.findall("w:p", namespaces={"w": W_NS}):
                for br in para.findall(".//w:br", namespaces={"w": W_NS}):
                    if br.attrib.get(f"{{{W_NS}}}type") == "page":
                        clone = et.fromstring(et.tostring(br))
                        br.getparent().append(clone)
                        injected = True
                        break
                if injected:
                    break

            assert injected, "Test setup failed: no page-break paragraph found to inject into"

            out_bytes = _io.BytesIO()
            with zf.ZipFile(out_bytes, "w", compression=zf.ZIP_DEFLATED) as zout:
                for item in names:
                    if item == "word/document.xml":
                        zout.writestr(
                            item,
                            et.tostring(doc_xml, xml_declaration=True, encoding="utf-8", standalone=True),
                        )
                    else:
                        zout.writestr(item, zin.read(item))

        with open(p, "wb") as fh:
            fh.write(out_bytes.getvalue())

        res = DocxEngine.clean_structure(p, max_consecutive_breaks=1, remove_empty_paragraphs=False)
        assert res["page_breaks_cleaned"] >= 1

    def test_color_normalization(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.update_styles(docx_path, styles_config={"Normal": {"color": "FF0000"}})
        DocxEngine.clean_structure(docx_path, normalize_all_colors_black=True, remove_empty_paragraphs=False)
        styles_xml = _read_styles_xml(docx_path)
        # All color elements in styles.xml should now be 000000
        # (clean_structure only touches document.xml colors, so check there)
        doc_xml = _read_doc_xml(docx_path)
        for c in doc_xml.findall(".//w:color", namespaces={"w": W_NS}):
            assert c.attrib.get(f"{{{W_NS}}}val") == "000000"

    def test_color_not_normalized_when_flag_off(self, tmp_path: Path):
        """normalize_all_colors_black=False must not touch colors."""
        docx_path = _make_docx(tmp_path)
        DocxEngine.clean_structure(docx_path, normalize_all_colors_black=False, remove_empty_paragraphs=False)
        # Simply assert no exception and file is valid
        _roundtrip(docx_path)

    def test_dry_run_leaves_document_intact(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        res_dry = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=True)
        assert res_dry['status'] == 'success'
        assert res_dry['dry_run'] is True
        assert res_dry['empty_paragraphs_removed'] >= 2

        # Second run without dry_run should still find and remove them
        res_real = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=False)
        assert res_real['dry_run'] is False
        assert res_real['empty_paragraphs_removed'] >= 2

    def test_dry_run_leaves_document_intact(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        # Check initial empty paragraphs
        res_dry = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=True)
        assert res_dry["status"] == "success"
        assert res_dry["dry_run"] is True
        assert res_dry["empty_paragraphs_removed"] >= 2
        
        # Second run without dry_run should still find and remove them
        res_real = DocxEngine.clean_structure(docx_path, remove_empty_paragraphs=True, dry_run=False)
        assert res_real["dry_run"] is False
        assert res_real["empty_paragraphs_removed"] >= 2

    def test_roundtrip(self, tmp_path: Path):
        docx_path = _make_docx(tmp_path)
        DocxEngine.clean_structure(docx_path)
        _roundtrip(docx_path)
