"""Gate 2 Verification: Full Document Styling Pipeline Parity Test.

Runs academic-id.yaml pipeline on identical input documents:
1. Headless backend (DocxEngine)
2. Live Word backend (LiveWordEngine via COM)
Compares resulting parameters and verifies parity across layout, styles, table formatting, and hygiene.
"""

import os
import shutil
import tempfile
from pathlib import Path
import yaml
import docx
import win32com.client
import pytest
from src.oxml_engine import DocxEngine
from src.live_engine import LiveWordEngine, com_worker, HAS_COM


pytestmark = pytest.mark.live


def create_sample_doc(target_path: Path):
    """Create sample unstyled document containing headings, paragraphs, empty paragraphs, and a table."""
    actions = [
        {"type": "heading", "text": "BAB I PENDAHULUAN", "level": 1},
        {"type": "paragraph", "text": "Latar belakang penelitian ini menjelaskan dasar teori dan urgensi."},
        {"type": "paragraph", "text": ""},  # empty paragraph to be cleaned
        {"type": "heading", "text": "1.1 Rumusan Masalah", "level": 2},
        {"type": "paragraph", "text": "Berikut adalah rumusan masalah yang akan dibahas dalam bab ini."},
        {
            "type": "table",
            "headers": ["No", "Parameter", "Nilai"],
            "rows": [["1", "Suhu", "25 C"], ["2", "Tekanan", "1 atm"]],
        },
        {"type": "paragraph", "text": ""},  # another empty paragraph
        {"type": "paragraph", "text": "Kesimpulan sementara dari pengujian awal."},
    ]
    DocxEngine.compose_document(str(target_path), actions, create_if_missing=True)


def run_pipeline_headless(file_path: str, preset: dict) -> dict:
    """Run academic-id pipeline via headless DocxEngine."""
    layout = preset.get("layout", {})
    res_layout = DocxEngine.set_page_layout(
        file_path,
        top_cm=layout.get("top_cm", 2.54),
        bottom_cm=layout.get("bottom_cm", 2.54),
        left_cm=layout.get("left_cm", 2.54),
        right_cm=layout.get("right_cm", 2.54),
        page_size=layout.get("page_size", "a4"),
        orientation=layout.get("orientation", "portrait"),
    )

    styles = preset.get("styles", {})
    res_styles = DocxEngine.update_styles(file_path, styles_config=styles)

    tbl_cfg = preset.get("table", {})
    res_table = DocxEngine.format_table_grid(
        file_path,
        border_sz=tbl_cfg.get("border_sz", 4),
        border_color=tbl_cfg.get("border_color", "000000"),
        header_shading=tbl_cfg.get("header_shading", "F2F2F2"),
        cant_split=tbl_cfg.get("cant_split", True),
        repeat_header=tbl_cfg.get("repeat_header", True),
        font_name=tbl_cfg.get("font_name", "Times New Roman"),
        font_size_pt=tbl_cfg.get("font_size_pt", 11.0),
    )

    hygiene = preset.get("hygiene", {})
    res_hygiene = DocxEngine.clean_structure(
        file_path,
        remove_empty_paragraphs=hygiene.get("remove_empty_paragraphs", True),
        remove_empty_tables=hygiene.get("remove_empty_tables", True),
        extract_single_image_tables=hygiene.get("extract_single_image_tables", True),
        normalize_all_colors_black=hygiene.get("normalize_all_colors_black", False),
    )

    return {
        "layout": res_layout,
        "styles": res_styles,
        "table": res_table,
        "hygiene": res_hygiene,
    }


def run_pipeline_live(file_path: str, preset: dict) -> dict:
    """Run academic-id pipeline via live LiveWordEngine attached to Word COM."""
    def _execute():
        app = win32com.client.GetActiveObject("Word.Application")
        wdoc = app.Documents.Open(str(file_path))
        try:
            layout = preset.get("layout", {})
            res_layout = LiveWordEngine.set_page_layout(
                wdoc,
                file_path,
                top_cm=layout.get("top_cm", 2.54),
                bottom_cm=layout.get("bottom_cm", 2.54),
                left_cm=layout.get("left_cm", 2.54),
                right_cm=layout.get("right_cm", 2.54),
                page_size=layout.get("page_size", "a4"),
                orientation=layout.get("orientation", "portrait"),
            )

            styles = preset.get("styles", {})
            res_styles = LiveWordEngine.update_styles(wdoc, file_path, styles_config=styles)

            tbl_cfg = preset.get("table", {})
            res_table = LiveWordEngine.format_table_grid(
                wdoc,
                file_path,
                border_sz=tbl_cfg.get("border_sz", 4),
                border_color=tbl_cfg.get("border_color", "000000"),
                header_shading=tbl_cfg.get("header_shading", "F2F2F2"),
                cant_split=tbl_cfg.get("cant_split", True),
                repeat_header=tbl_cfg.get("repeat_header", True),
                font_name=tbl_cfg.get("font_name", "Times New Roman"),
                font_size_pt=tbl_cfg.get("font_size_pt", 11.0),
            )

            hygiene = preset.get("hygiene", {})
            res_hygiene = LiveWordEngine.clean_structure(
                wdoc,
                file_path,
                remove_empty_paragraphs=hygiene.get("remove_empty_paragraphs", True),
                remove_empty_tables=hygiene.get("remove_empty_tables", True),
                extract_single_image_tables=hygiene.get("extract_single_image_tables", True),
                normalize_all_colors_black=hygiene.get("normalize_all_colors_black", False),
            )

            # Read final state from Word
            outline = LiveWordEngine.read_document(wdoc, file_path, mode="outline")
            tbl_info = LiveWordEngine.read_document(wdoc, file_path, mode="table", table_index=0)

            return {
                "layout": res_layout,
                "styles": res_styles,
                "table": res_table,
                "hygiene": res_hygiene,
                "outline": outline,
                "table_info": tbl_info,
            }
        finally:
            wdoc.Close(SaveChanges=False)

    return com_worker.run(_execute)


def test_academic_id_pipeline_parity(tmp_path: Path):
    if not HAS_COM:
        pytest.skip("pywin32 not available")

    preset_path = Path(".agents/presets/academic-id.yaml")
    with open(preset_path, "r", encoding="utf-8") as f:
        preset = yaml.safe_load(f)

    headless_doc = tmp_path / "doc_headless.docx"
    live_doc = tmp_path / "doc_live.docx"

    create_sample_doc(headless_doc)
    create_sample_doc(live_doc)

    out_headless = run_pipeline_headless(str(headless_doc), preset)
    out_live = run_pipeline_live(str(live_doc), preset)

    # 1. Layout Parity
    assert out_headless["layout"]["page_size"] == out_live["layout"]["page_size"] == "a4"
    assert out_headless["layout"]["orientation"] == out_live["layout"]["orientation"] == "portrait"
    assert out_headless["layout"]["margins_cm"] == out_live["layout"]["margins_cm"]

    # 2. Styles Parity
    assert "Normal" in out_live["styles"]["updated_styles"]
    assert "Heading 1" in out_live["styles"]["updated_styles"]
    assert "Heading 2" in out_live["styles"]["updated_styles"]

    # 3. Table Formatting Parity
    assert out_headless["table"]["tables_formatted"] == out_live["table"]["tables_formatted"] == 1
    assert out_live["table_info"]["rows_count"] == 3
    assert out_live["table_info"]["columns_count"] == 3

    # 4. Hygiene Parity
    assert out_headless["hygiene"]["empty_paragraphs_removed"] == 2
    assert out_live["hygiene"]["empty_paragraphs_removed"] == 2



