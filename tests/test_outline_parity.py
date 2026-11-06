"""Permanent verification test for outline parity between Headless and Live engines."""

import time
from pathlib import Path
from typing import Any, Generator
import pytest
import win32com.client

from src.oxml_engine import DocxEngine
from src.live_engine import LiveWordEngine, com_worker, HAS_COM


pytestmark = pytest.mark.live


def _is_word_active() -> bool:
    if not HAS_COM:
        return False
    try:
        def _check():
            try:
                app = win32com.client.GetActiveObject("Word.Application")
                return app is not None
            except Exception:
                return False
        return bool(com_worker.run(_check, timeout=3.0))
    except Exception:
        return False


def _clean_lines(text: str) -> list[str]:
    """Strip backend prefix and trailing whitespaces for line-by-line comparison."""
    return [line.rstrip() for line in text.splitlines() if not line.startswith("[backend:")]


def test_outline_parity_short_document(tmp_path: Path):
    """Verify outline parity on documents under 40 lines (headings, body, tables)."""
    if not _is_word_active():
        pytest.skip("Microsoft Word is not active.")

    test_file = (tmp_path / "short_doc.docx").resolve()
    actions = [
        {"type": "heading", "text": "Executive Summary", "level": 1},
        {"type": "paragraph", "text": "This is introductory body paragraph 1."},
        {"type": "paragraph", "text": "This is body paragraph 2 between headings."},
        {"type": "heading", "text": "Technical Specifications", "level": 2},
        {"type": "paragraph", "text": "Detailed specs description."},
        {
            "type": "table",
            "headers": ["Param", "Specification", "Status"],
            "rows": [["Memory", "16GB", "Passed"], ["Disk", "512GB", "Passed"]],
        },
        {"type": "paragraph", "text": "Final concluding remarks after table."},
    ]
    DocxEngine.compose_document(str(test_file), actions, create_if_missing=True)

    headless_out = DocxEngine.read_document(str(test_file), mode="outline")

    def _live():
        app = win32com.client.GetActiveObject("Word.Application")
        wdoc = app.Documents.Open(str(test_file))
        try:
            return LiveWordEngine.read_document(wdoc, str(test_file), mode="outline")
        finally:
            wdoc.Close(SaveChanges=False)

    live_out = com_worker.run(_live)

    h_lines = _clean_lines(headless_out)
    l_lines = _clean_lines(live_out)

    assert h_lines == l_lines, f"Outline mismatch on short document:\nHeadless:\n{headless_out}\nLive:\n{live_out}"


def test_outline_parity_large_document(tmp_path: Path):
    """Verify outline parity on 40+ lines document with headings and body paragraphs."""
    if not _is_word_active():
        pytest.skip("Microsoft Word is not active.")

    test_file = (tmp_path / "large_doc.docx").resolve()
    actions = []
    for sec_i in range(1, 15):
        actions.append({"type": "heading", "text": f"Chapter {sec_i}: Domain Architecture", "level": 1})
        for p_i in range(1, 4):
            actions.append({"type": "paragraph", "text": f"Chapter {sec_i} paragraph {p_i} containing standard body text."})
        if sec_i % 3 == 0:
            actions.append({
                "type": "table",
                "headers": [f"Ch{sec_i}_Col1", f"Ch{sec_i}_Col2"],
                "rows": [["Val A", "Val B"]],
            })

    DocxEngine.compose_document(str(test_file), actions, create_if_missing=True)

    headless_out = DocxEngine.read_document(str(test_file), mode="outline")

    def _live():
        app = win32com.client.GetActiveObject("Word.Application")
        wdoc = app.Documents.Open(str(test_file))
        try:
            return LiveWordEngine.read_document(wdoc, str(test_file), mode="outline")
        finally:
            wdoc.Close(SaveChanges=False)

    live_out = com_worker.run(_live)

    h_lines = _clean_lines(headless_out)
    l_lines = _clean_lines(live_out)

    assert h_lines == l_lines, f"Outline mismatch on large document:\nHeadless:\n{headless_out}\nLive:\n{live_out}"


def test_outline_parity_tracked_deletion(tmp_path: Path):
    """Verify that a paragraph deleted under Track Changes is excluded from [p:N] outline."""
    if not _is_word_active():
        pytest.skip("Microsoft Word is not active.")

    test_file = (tmp_path / "del_doc.docx").resolve()
    actions = [
        {"type": "heading", "text": "Document Title", "level": 1},
        {"type": "paragraph", "text": "Paragraph 1 to keep."},
        {"type": "paragraph", "text": "Paragraph 2 to delete under track changes."},
        {"type": "paragraph", "text": "Paragraph 3 to keep."},
    ]
    DocxEngine.compose_document(str(test_file), actions, create_if_missing=True)

    def _delete_and_read():
        app = win32com.client.GetActiveObject("Word.Application")
        wdoc = app.Documents.Open(str(test_file))
        try:
            # Delete paragraph [p:2] (which is COM Paragraphs.Item(3))
            LiveWordEngine.mutate_element(wdoc, str(test_file), target_index=2, action="delete")
            live_outline = LiveWordEngine.read_document(wdoc, str(test_file), mode="outline")
            return live_outline
        finally:
            wdoc.Close(SaveChanges=False)

    live_out = com_worker.run(_delete_and_read)

    assert "Paragraph 2 to delete" not in live_out
    assert "[p:0]" in live_out
    assert "[p:1]" in live_out
    assert "[p:2]" in live_out
    assert "Paragraph 3 to keep" in live_out


def test_outline_performance_1000_paragraphs(tmp_path: Path):
    """Measure outline extraction performance across 1,000 paragraphs against the 2.0s budget."""
    if not _is_word_active():
        pytest.skip("Microsoft Word is not active.")

    test_file = (tmp_path / "perf_doc.docx").resolve()
    actions = []
    for h_i in range(1, 21):
        actions.append({"type": "heading", "text": f"Major Section {h_i}", "level": 1})
        for p_i in range(1, 50):
            actions.append({"type": "paragraph", "text": f"Body paragraph {p_i} in section {h_i} describing system metrics."})
        if h_i % 5 == 0:
            actions.append({
                "type": "table",
                "headers": ["Metric", "Value"],
                "rows": [["Throughput", "1000/s"], ["Latency", "5ms"]],
            })

    DocxEngine.compose_document(str(test_file), actions, create_if_missing=True)

    def _measure():
        app = win32com.client.GetActiveObject("Word.Application")
        wdoc = app.Documents.Open(str(test_file))
        try:
            t0 = time.perf_counter()
            outline = LiveWordEngine.read_document(wdoc, str(test_file), mode="outline")
            elapsed = time.perf_counter() - t0
            return elapsed, outline
        finally:
            wdoc.Close(SaveChanges=False)

    elapsed, outline = com_worker.run(_measure)
    print(f"\n1,000-paragraph outline time: {elapsed:.3f}s")
    assert elapsed < 2.0, f"Outline extraction exceeded 2.0s budget: {elapsed:.3f}s"
    assert "Major Section 1" in outline
    assert "Major Section 20" in outline

