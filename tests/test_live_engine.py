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
    """Create documents/live_test.docx, open in running Word, and close at teardown."""
    if not _is_word_active():
        pytest.skip("Microsoft Word desktop instance is not currently running via COM.")

    docs_dir = Path("documents")
    docs_dir.mkdir(exist_ok=True)
    test_file = (docs_dir / "live_test.docx").resolve()

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
