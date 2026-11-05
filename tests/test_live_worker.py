"""Unit tests for COM Worker thread, timeouts, poisoned state, and routing (no Word required)."""

from __future__ import annotations

import time
import pytest
from pathlib import Path
from src.live_engine import (
    ComWorker,
    LiveWordError,
    ERROR_MESSAGES,
    _norm,
    paths_match,
    has_word_lock_file,
    pick_backend,
)


class TestComWorker:
    def test_basic_execution(self):
        worker = ComWorker()
        try:
            res = worker.run(lambda x, y: x + y, 10, 25, timeout=2.0)
            assert res == 35
        finally:
            worker.stop()

    def test_timeout_and_poisoned_state_and_recovery(self):
        worker = ComWorker()
        try:
            # Task that takes 0.6 seconds, but timeout is 0.1s
            def slow_task():
                time.sleep(0.5)
                return "done_slow"

            with pytest.raises(LiveWordError) as exc_info:
                worker.run(slow_task, timeout=0.1)

            assert "Word may be waiting on a modal dialog" in str(exc_info.value)
            assert worker.is_poisoned() is True

            # Subsequent call fails immediately with poisoned error without waiting
            t0 = time.time()
            with pytest.raises(LiveWordError) as exc_info2:
                worker.run(lambda: "fast", timeout=1.0)
            elapsed = time.time() - t0
            assert elapsed < 0.05
            assert "earlier Word operation has not returned" in str(exc_info2.value)

            # Wait for slow_task to complete in worker thread
            time.sleep(0.6)
            assert worker.is_poisoned() is False

            # Subsequent call now succeeds normally
            res = worker.run(lambda: "recovered", timeout=1.0)
            assert res == "recovered"
        finally:
            worker.stop()

    def test_single_thread_preserved(self):
        worker = ComWorker()
        try:
            tid1 = worker.run(lambda: worker._thread.ident if worker._thread else None, timeout=1.0)
            tid2 = worker.run(lambda: worker._thread.ident if worker._thread else None, timeout=1.0)
            assert tid1 is not None
            assert tid1 == tid2
        finally:
            worker.stop()

    def test_exception_propagation(self):
        worker = ComWorker()
        try:
            def failing():
                raise ValueError("test error inside worker")

            with pytest.raises(ValueError, match="test error inside worker"):
                worker.run(failing, timeout=1.0)
        finally:
            worker.stop()


class TestPathAndLockUtilities:
    def test_paths_match_same_file(self, tmp_path: Path):
        f = tmp_path / "doc.docx"
        f.write_text("dummy")
        assert paths_match(str(f), str(f))
        assert paths_match(str(f).upper(), str(f).lower())

    def test_has_word_lock_file(self, tmp_path: Path):
        doc = tmp_path / "report.docx"
        doc.write_text("dummy")
        assert has_word_lock_file(doc) is False

        # Create standard Word owner lock file: ~$report.docx or ~$port.docx
        lock = tmp_path / "~$port.docx"
        lock.write_text("lock")
        assert has_word_lock_file(doc) is True


class TestPickBackendNoWord:
    def test_pick_backend_headless_fallback(self, tmp_path: Path):
        doc = tmp_path / "test.docx"
        backend, com_doc = pick_backend(doc, tool_name="test", is_mutation=False)
        assert backend == "headless"
        assert com_doc is None

    def test_pick_backend_forced_live_raises_when_no_word(self, tmp_path: Path, monkeypatch):
        monkeypatch.setenv("WORD_BACKEND", "live")
        doc = tmp_path / "test.docx"
        with pytest.raises(LiveWordError, match="WORD_BACKEND is set to 'live'"):
            pick_backend(doc, tool_name="test", is_mutation=False)

    def test_lock_file_blocks_mutating_when_no_word_attached(self, tmp_path: Path):
        doc = tmp_path / "locked.docx"
        doc.write_text("data")
        lock = tmp_path / "~$cked.docx"
        lock.write_text("owner")

        # Mutating tool should raise lock error
        with pytest.raises(LiveWordError, match="lock file found"):
            pick_backend(doc, tool_name="mutate_element", is_mutation=True)

        # Reading tool should fall back to headless
        backend, _ = pick_backend(doc, tool_name="read_document", is_mutation=False)
        assert backend == "headless"
