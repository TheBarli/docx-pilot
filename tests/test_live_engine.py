"""Automated live tests for Microsoft Word COM backend (Phase 1).

Requires Microsoft Word to be actively running on Windows desktop.
Auto-skips gracefully if Word is not running.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Generator
import pytest

from src.oxml_engine import DocxEngine
from src.live_engine import (
    HAS_COM,
    LiveWordEngine,
    com_worker,
    find_open_document,
    pick_backend,
    paths_match,
    LiveWordError,
)

if HAS_COM:
    import win32com.client


def _is_word_active() -> bool:
    """Check if Microsoft Word is active and accessible via COM."""
    if not HAS_COM:
        return False
    def _check():
        try:
            app = win32com.client.GetActiveObject("Word.Application")
            return app is not None
        except Exception:
            return False
    try:
        return bool(com_worker.run(_check, timeout=3.0))
    except Exception:
        return False


pytestmark = pytest.mark.live


@pytest.fixture
def live_doc(tmp_path: Path) -> Generator[tuple[Path, Any], None, None]:
    """Create isolated test doc, open in running Word, and close at teardown."""
    if not _is_word_active():
        pytest.skip("Microsoft Word desktop instance is not currently running via COM.")

    test_file = (tmp_path / "live_test.docx").resolve()

    # Initial composition via DocxEngine (headless)
    actions = [
        {"type": "heading", "text": "Live Test Document", "level": 1},
        {"type": "paragraph", "text": "First body paragraph for testing live COM engine."},
        {"type": "paragraph", "text": "Second body paragraph with target keyword."},
        {
            "type": "table",
            "headers": ["Metric", "Value"],
            "rows": [["Memory", "64MB"], ["Latency", "12ms"]],
        },
        {"type": "paragraph", "text": "Third body paragraph after table."},
    ]
    DocxEngine.compose_document(str(test_file), actions, create_if_missing=True)

    # Open in Word via COM worker
    def _open():
        app = win32com.client.GetActiveObject("Word.Application")
        # Documents.Open without activating or altering views
        doc = app.Documents.Open(str(test_file))
        return doc

    doc = com_worker.run(_open)

    try:
        yield test_file, doc
    finally:
        # Close without saving changes (C12)
        def _close():
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass

        com_worker.run(_close)
        if test_file.exists():
            try:
                test_file.unlink()
            except Exception:
                pass


def test_detection_open_and_closed(live_doc):
    test_file, com_doc = live_doc

    # Detection when open
    found = find_open_document(test_file)
    assert found is not None
    # Access COM property inside worker thread per Rule C9
    found_full_name = com_worker.run(lambda: found.FullName)
    assert paths_match(found_full_name, test_file)

    backend, doc_ref = pick_backend(test_file, tool_name="read_document", is_mutation=False)
    assert backend == "live"
    assert doc_ref is not None

    # Path variant matching (uppercase vs lowercase)
    upper_path = str(test_file).upper()
    assert paths_match(upper_path, found_full_name)


def test_live_read_outline_and_full(live_doc):
    test_file, com_doc = live_doc

    outline = LiveWordEngine.read_document(com_doc, test_file, mode="outline")
    assert isinstance(outline, str)
    assert "# Outline: live_test.docx" in outline
    assert "- `[p:0]` **(Heading 1)**: \"Live Test Document\"" in outline
    assert "`[table:0]`" in outline

    full = LiveWordEngine.read_document(com_doc, test_file, mode="full")
    assert isinstance(full, str)
    assert "`[p:0]` (Heading 1): Live Test Document" in full
    assert "`[p:1]`" in full
    assert "`[table:0]` Table" in full


def test_live_mutations_and_undo(live_doc):
    test_file, com_doc = live_doc

    # 1. Insert before [p:1]
    res1 = LiveWordEngine.mutate_element(
        com_doc,
        test_file,
        target_index=1,
        action="insert_before",
        content={"text": "Inserted before p1", "style": "Normal"},
    )
    assert res1["status"] == "success"
    assert "local_outline" in res1

    # Verify paragraph inserted
    full1 = LiveWordEngine.read_document(com_doc, test_file, mode="full")
    assert "Inserted before p1" in full1

    # Undo record test (C5): exactly one undo reverts the action
    def _undo():
        com_doc.Undo()

    com_worker.run(_undo)

    full_after_undo = LiveWordEngine.read_document(com_doc, test_file, mode="full")
    assert "Inserted before p1" not in full_after_undo


def test_live_replace_and_delete(live_doc):
    test_file, com_doc = live_doc

    # Replace [p:1]
    res_replace = LiveWordEngine.mutate_element(
        com_doc,
        test_file,
        target_index=1,
        action="replace",
        content={"text": "Replaced Body Paragraph 1", "bold": True},
    )
    assert res_replace["status"] == "success"
    assert res_replace["updated_text"] == "Replaced Body Paragraph 1"

    full = LiveWordEngine.read_document(com_doc, test_file, mode="full")
    assert "Replaced Body Paragraph 1" in full

    # Delete [p:1]
    res_del = LiveWordEngine.mutate_element(
        com_doc,
        test_file,
        target_index=1,
        action="delete",
    )
    assert res_del["status"] == "success"

    full_after_del = LiveWordEngine.read_document(com_doc, test_file, mode="full")
    assert "Replaced Body Paragraph 1" not in full_after_del


def test_live_find_and_replace(live_doc):
    test_file, com_doc = live_doc

    res = LiveWordEngine.find_and_replace(
        com_doc,
        test_file,
        find_text="target keyword",
        replace_text="SUPER_SUCCESSFUL_REPLACEMENT",
        match_case=True,
    )
    assert res["status"] == "success"
    assert res["matches_replaced"] == 1

    full = LiveWordEngine.read_document(com_doc, test_file, mode="full")
    assert "SUPER_SUCCESSFUL_REPLACEMENT" in full


def test_find_and_replace_max_length_validation(live_doc):
    test_file, com_doc = live_doc
    with pytest.raises(ValueError, match="exceeds Word 255 character limit"):
        LiveWordEngine.find_and_replace(
            com_doc,
            test_file,
            find_text="x" * 256,
            replace_text="replacement",
        )


def test_live_read_table(live_doc):
    test_file, com_doc = live_doc
    res = LiveWordEngine.read_document(com_doc, test_file, mode="table", table_index=0)
    assert isinstance(res, dict)
    assert res["table_index"] == 0
    assert res["rows_count"] == 3
    assert res["columns_count"] == 2
    assert res["matrix"][0] == ["Metric", "Value"]
    assert res["matrix"][1] == ["Memory", "64MB"]
    assert "| Metric | Value |" in res["markdown"]


def test_live_set_page_layout(live_doc):
    test_file, com_doc = live_doc
    res = LiveWordEngine.set_page_layout(
        com_doc,
        test_file,
        top_cm=3.0,
        bottom_cm=3.0,
        left_cm=2.0,
        right_cm=2.0,
        page_size="a4",
        orientation="landscape",
    )
    assert res["status"] == "success"
    assert res["page_size"] == "a4"
    assert res["orientation"] == "landscape"

    # Verify directly via COM worker
    def _verify():
        ps = com_doc.Sections.Item(1).PageSetup
        return ps.Orientation, round(ps.TopMargin)

    orient, top_pt = com_worker.run(_verify)
    assert orient == 1  # wdOrientLandscape
    assert abs(top_pt - round(3.0 * 72.0 / 2.54)) <= 1


def test_live_update_styles(live_doc):
    test_file, com_doc = live_doc
    res = LiveWordEngine.update_styles(
        com_doc,
        test_file,
        styles_config={
            "Normal": {
                "font": "Arial",
                "size_pt": 12.0,
                "bold": True,
                "color": "FF0000",
                "space_after_pt": 6.0,
            }
        },
    )
    assert res["status"] == "success"
    assert "Normal" in res["updated_styles"]

    def _verify():
        st = com_doc.Styles("Normal")
        return st.Font.Name, st.Font.Size, st.Font.Bold

    name, size, bold = com_worker.run(_verify)
    assert name == "Arial"
    assert size == 12.0
    assert bold == -1 or bold is True  # wdToggle / True in COM


def test_live_compose_document(live_doc):
    test_file, com_doc = live_doc
    actions = [
        {"type": "heading", "text": "Appended Live Heading", "level": 2},
        {"type": "paragraph", "text": "Appended Live Body text", "bold": False},
    ]
    res = LiveWordEngine.compose_document(com_doc, test_file, actions=actions)
    assert res["status"] == "success"
    assert res["total_actions"] == 2

    full = LiveWordEngine.read_document(com_doc, test_file, mode="full")
    assert "Appended Live Heading" in full
    assert "Appended Live Body text" in full


def test_live_format_table_grid(live_doc):
    test_file, com_doc = live_doc
    res = LiveWordEngine.format_table_grid(
        com_doc,
        test_file,
        table_indices=[0],
        border_sz=4,
        border_color="000000",
        header_shading="E0E0E0",
        cant_split=True,
        repeat_header=True,
    )
    assert res["status"] == "success"
    assert res["tables_formatted"] == 1

    def _verify():
        tbl = com_doc.Tables.Item(1)
        r1 = tbl.Rows.Item(1)
        return r1.HeadingFormat, tbl.Rows.AllowBreakAcrossPages

    h_format, allow_break = com_worker.run(_verify)
    assert h_format == -1  # True
    assert allow_break == 0  # False


def test_live_clean_structure_dry_run_and_execution(live_doc):
    test_file, com_doc = live_doc

    # Add an empty paragraph to the document
    def _add_empty():
        end_rng = com_doc.Range(com_doc.Content.End - 1, com_doc.Content.End - 1)
        p = com_doc.Paragraphs.Add(Range=end_rng)
        p.Range.Text = "\n"

    com_worker.run(_add_empty)

    # 1. Test dry_run: must report >= 1 empty paragraph and not mutate
    res_dry = LiveWordEngine.clean_structure(
        com_doc,
        test_file,
        remove_empty_paragraphs=True,
        dry_run=True,
    )
    assert res_dry["status"] == "success"
    assert res_dry["dry_run"] is True
    assert res_dry["empty_paragraphs_removed"] >= 1

    # 2. Test actual removal: must remove it
    res_real = LiveWordEngine.clean_structure(
        com_doc,
        test_file,
        remove_empty_paragraphs=True,
        dry_run=False,
    )
    assert res_real["status"] == "success"
    assert res_real["dry_run"] is False
    assert res_real["empty_paragraphs_removed"] >= 1

