"""Live Microsoft Word backend engine via COM for MsWord-ai-integration.

Provides detection of open Word documents, routing between live and headless backends,
a single-threaded COM worker with timeouts and poisoned-state protection,
and undo/track revision wrappers.
"""

from __future__ import annotations

import os
import queue
import sys
import threading
from concurrent.futures import Future
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable, Dict, Generator, List, Optional, Tuple, Union

# ---------------------------------------------------------------------------
# Guarded COM Imports (C2: pywin32 is optional)
# ---------------------------------------------------------------------------
HAS_COM = False
try:
    if sys.platform == "win32":
        import pythoncom
        import win32api
        import win32com.client
        HAS_COM = True
except ImportError:
    HAS_COM = False


# ---------------------------------------------------------------------------
# Custom Exceptions & Standardized Error Catalog (Section 3.4)
# ---------------------------------------------------------------------------
class LiveWordError(Exception):
    """Base exception for all Live Word COM interactions with user-friendly messages."""
    pass


ERROR_MESSAGES: Dict[str, str] = {
    "index_out_of_range": (
        "Index [p:{target_index}] is out of range ({valid_range})."
    ),
    "read_only_or_protected": (
        "The document '{file_name}' is read-only or protected ({state}). "
        "Please enable editing in Microsoft Word."
    ),
    "protected_view": (
        "The document '{file_name}' is open in Protected View. "
        "Please click 'Enable Editing' in Microsoft Word."
    ),
    "cloud_hosted": (
        "The document '{file_name}' is cloud-hosted on OneDrive/SharePoint ({url}). "
        "Please save or copy it locally, or close it in Word and set WORD_BACKEND=headless."
    ),
    "modal_dialog_or_timeout": (
        "Word may be waiting on a modal dialog or is not responding "
        "(call timed out after {timeout}s). Please check Microsoft Word and dismiss any open dialogs."
    ),
    "worker_poisoned": (
        "An earlier Word operation has not returned; Word is probably still blocked by a modal dialog. "
        "Please check Microsoft Word and dismiss any open dialogs."
    ),
    "not_open_in_word": (
        "WORD_BACKEND is set to 'live', but the document '{file_name}' is not currently open in Microsoft Word."
    ),
    "lock_file_cannot_attach": (
        "The document '{file_name}' appears to be open in Word (lock file found), "
        "but could not be attached via COM. Word may be running elevated (as Administrator), "
        "in a different session, or the lock file is stale. Please close Word or verify the lock."
    ),
    "tool_not_yet_live_capable": (
        "Tool '{tool_name}' is not yet live-capable in this phase. "
        "The file is open in Microsoft Word, so no changes were written to disk."
    ),
}


# ---------------------------------------------------------------------------
# Path Normalization & Matching (Section 3.2)
# ---------------------------------------------------------------------------
def _norm(file_path: Union[str, Path]) -> str:
    """Normalize path: realpath, normcase, and expand 8.3 short names if on Windows."""
    p_str = str(file_path)
    try:
        real_p = os.path.realpath(p_str)
    except Exception:
        real_p = p_str

    norm_p = os.path.normcase(real_p)
    if HAS_COM and sys.platform == "win32" and os.path.exists(norm_p):
        try:
            expanded = win32api.GetLongPathName(norm_p)
            norm_p = os.path.normcase(expanded)
        except Exception:
            pass
    return norm_p


def paths_match(path1: Union[str, Path], path2: Union[str, Path]) -> bool:
    """Compare two paths accurately across case, short names, and mapped drives."""
    n1 = _norm(path1)
    n2 = _norm(path2)
    if n1 == n2:
        return True

    s1 = str(path1)
    s2 = str(path2)
    if os.path.exists(s1) and os.path.exists(s2):
        try:
            return os.path.samefile(s1, s2)
        except Exception:
            pass
    return False


def has_word_lock_file(file_path: Union[str, Path]) -> bool:
    """Check if Word has generated a temporary owner lock file (~$...) in the directory."""
    try:
        p = Path(file_path).resolve()
        parent = p.parent
        if not parent.exists():
            return False
        name = p.name
        candidates = [f"~${name}"]
        if len(name) > 2:
            candidates.append(f"~${name[2:]}")
        if len(name) > 1:
            candidates.append(f"~${name[1:]}")
        for c in candidates:
            if (parent / c).exists():
                return True
    except Exception:
        pass
    return False


# ---------------------------------------------------------------------------
# COM Worker Thread with Timeout & Poisoned State Protection (Section 3.8, C9)
# ---------------------------------------------------------------------------
class ComWorker:
    """Dedicated single worker thread for COM operations with timeout & poisoned state."""

    def __init__(self, default_timeout: float = 20.0):
        self._default_timeout = float(os.getenv("LIVE_CALL_TIMEOUT", default_timeout))
        self._queue: queue.Queue[Optional[Tuple[Callable[..., Any], tuple, dict, Future]]] = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._poisoned = False
        self._lock = threading.Lock()
        self._running = False

    def start(self) -> None:
        """Start worker thread if not already active."""
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._running = True
                self._thread = threading.Thread(
                    target=self._worker_loop,
                    name="LiveWordCOMWorker",
                    daemon=True,
                )
                self._thread.start()

    def _worker_loop(self) -> None:
        """Thread worker loop running CoInitialize/CoUninitialize."""
        if HAS_COM:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

        try:
            while self._running:
                item = self._queue.get()
                if item is None:
                    self._queue.task_done()
                    break

                func, args, kwargs, future = item
                try:
                    result = func(*args, **kwargs)
                    if not future.done():
                        future.set_result(result)
                except BaseException as exc:
                    if not future.done():
                        future.set_exception(exc)
                finally:
                    self._queue.task_done()
                    with self._lock:
                        # If a prior call timed out, its completion clears the poisoned state
                        if self._poisoned:
                            self._poisoned = False
        finally:
            if HAS_COM:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def run(self, func: Callable[..., Any], *args: Any, timeout: Optional[float] = None, **kwargs: Any) -> Any:
        """Submit a callable to the COM worker thread and await its result within timeout."""
        self.start()
        effective_timeout = timeout if timeout is not None else self._default_timeout

        with self._lock:
            if self._poisoned:
                raise LiveWordError(ERROR_MESSAGES["worker_poisoned"])

        fut: Future = Future()
        self._queue.put((func, args, kwargs, fut))

        try:
            return fut.result(timeout=effective_timeout)
        except TimeoutError:
            with self._lock:
                self._poisoned = True
            msg = ERROR_MESSAGES["modal_dialog_or_timeout"].format(timeout=effective_timeout)
            raise LiveWordError(msg) from None

    def is_poisoned(self) -> bool:
        """Check if worker is currently in poisoned state."""
        with self._lock:
            return self._poisoned

    def clear_poison(self) -> None:
        """Manual reset of poisoned flag if needed."""
        with self._lock:
            self._poisoned = False

    def stop(self) -> None:
        """Stop the worker thread cleanly."""
        with self._lock:
            self._running = False
        self._queue.put(None)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=2.0)
        self._thread = None


# Global singleton COM worker instance
com_worker = ComWorker()


# ---------------------------------------------------------------------------
# Live Detection & Routing (Section 3.2)
# ---------------------------------------------------------------------------
def find_open_document(file_path: Union[str, Path]) -> Optional[Any]:
    """Search currently active Microsoft Word instance for an open document matching path.

    Returns the COM Document object if found, or None.
    Never launches Word (C4).
    """
    if not HAS_COM:
        return None

    def _query() -> Optional[Any]:
        try:
            # C4: Attach ONLY via GetActiveObject, never Dispatch or EnsureDispatch
            app = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return None

        # 1. Search in open Documents
        try:
            docs = app.Documents
            count = docs.Count
            for i in range(1, count + 1):
                try:
                    doc = docs.Item(i)
                    full_name = doc.FullName
                    if paths_match(file_path, full_name):
                        return doc
                except Exception:
                    continue
        except Exception:
            pass

        return None

    return com_worker.run(_query, timeout=5.0)


def pick_backend(
    file_path: Union[str, Path],
    tool_name: str = "tool",
    is_mutation: bool = False,
) -> Tuple[str, Optional[Any]]:
    """Determine whether to route request to 'live' or 'headless' backend (Section 3.2).

    Returns:
        tuple (backend_name, com_doc_or_None)
    Raises:
        LiveWordError under specified safety net conditions.
    """
    mode_env = os.getenv("WORD_BACKEND", "auto").lower().strip()
    target_path = Path(file_path)
    file_name = target_path.name

    # Step 1: Explicit headless override
    if mode_env == "headless":
        return "headless", None

    # Step 2: Check if COM available or Word running
    if not HAS_COM:
        if has_word_lock_file(target_path):
            if is_mutation:
                raise LiveWordError(ERROR_MESSAGES["lock_file_cannot_attach"].format(file_name=file_name))
        if mode_env == "live":
            raise LiveWordError(ERROR_MESSAGES["not_open_in_word"].format(file_name=file_name))
        return "headless", None

    # Try attaching to running Word
    def _inspect_word() -> Tuple[Optional[str], Optional[Any]]:
        try:
            app = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return None, None

        # Step 3: Match in Application.Documents
        try:
            docs = app.Documents
            for i in range(1, docs.Count + 1):
                try:
                    d = docs.Item(i)
                    if paths_match(target_path, d.FullName):
                        return "live", d
                except Exception:
                    continue
        except Exception:
            pass

        # Step 4: Check ProtectedViewWindows
        try:
            pv_windows = app.ProtectedViewWindows
            for i in range(1, pv_windows.Count + 1):
                try:
                    pv_doc = pv_windows.Item(i).Document
                    if paths_match(target_path, pv_doc.FullName):
                        raise LiveWordError(ERROR_MESSAGES["protected_view"].format(file_name=file_name))
                except LiveWordError:
                    raise
                except Exception:
                    continue
        except LiveWordError:
            raise
        except Exception:
            pass

        # Step 5: Check cloud-hosted documents (OneDrive/SharePoint URLs)
        try:
            docs = app.Documents
            for i in range(1, docs.Count + 1):
                try:
                    d = docs.Item(i)
                    fn = str(d.FullName).lower()
                    if fn.startswith("http://") or fn.startswith("https://"):
                        if str(d.Name).lower() == file_name.lower():
                            raise LiveWordError(
                                ERROR_MESSAGES["cloud_hosted"].format(file_name=file_name, url=d.FullName)
                            )
                except LiveWordError:
                    raise
                except Exception:
                    continue
        except LiveWordError:
            raise
        except Exception:
            pass

        return None, None

    status, doc = com_worker.run(_inspect_word, timeout=5.0)
    if status == "live" and doc is not None:
        return "live", doc

    # Step 6: Lock-file safety net
    if has_word_lock_file(target_path):
        if is_mutation:
            raise LiveWordError(ERROR_MESSAGES["lock_file_cannot_attach"].format(file_name=file_name))
        return "headless", None

    # Forced live check
    if mode_env == "live":
        raise LiveWordError(ERROR_MESSAGES["not_open_in_word"].format(file_name=file_name))

    # Step 7: Headless fallback
    return "headless", None


# ---------------------------------------------------------------------------
# Reading Semantics & 0-Based Index Mapping (Section 3.5, C1)
# ---------------------------------------------------------------------------
def _visible_map(doc: Any) -> Dict[str, Any]:
    """Build bulk map of 0-based visible body paragraphs excluding tables & deleted paragraphs.

    Stub signature for Phase 0; populated in Phase 1 (E1/3.5/3.11).
    """
    return {
        "indices": [],       # 1-based COM paragraph indices corresponding to visible body [p:N]
        "deleted_spans": {}, # per-paragraph deleted character ranges
    }


def _com_idx(doc: Any, n: int) -> int:
    """Map tool 0-based paragraph index n to COM 1-based paragraph index using _visible_map.

    Never inlines `+ 1` (C1).
    """
    vmap = _visible_map(doc)
    indices = vmap.get("indices", [])
    if n < 0 or n >= len(indices):
        valid = f"0..{len(indices) - 1}" if indices else "empty"
        raise LiveWordError(ERROR_MESSAGES["index_out_of_range"].format(target_index=n, valid_range=valid))
    return indices[n]


# ---------------------------------------------------------------------------
# Undo & Track Changes Wrappers (Section 3.6, C5)
# ---------------------------------------------------------------------------
@contextmanager
def tracking(doc: Any, on: bool) -> Generator[None, None, None]:
    """Temporarily flip TrackRevisions for a sub-step inside an existing undo record."""
    prev_track = None
    try:
        prev_track = doc.TrackRevisions
        doc.TrackRevisions = on
    except Exception:
        pass

    try:
        yield
    finally:
        if prev_track is not None:
            try:
                doc.TrackRevisions = prev_track
            except Exception:
                pass


@contextmanager
def live_call(
    doc: Any,
    tool_name: str,
    target_desc: str,
    *,
    track: Optional[bool] = None,
) -> Generator[None, None, None]:
    """Context manager for all mutating live COM calls ensuring exactly one undo step (C5)."""
    undo_record = None
    prev_track = None
    try:
        app = doc.Application
        undo_record = app.UndoRecord
        undo_record.StartCustomRecord(f"Agent: {tool_name} {target_desc}")
    except Exception:
        undo_record = None

    try:
        prev_track = doc.TrackRevisions
        if track is True:
            doc.TrackRevisions = True
    except Exception:
        pass

    try:
        yield
    finally:
        if prev_track is not None:
            try:
                doc.TrackRevisions = prev_track
            except Exception:
                pass
        if undo_record is not None:
            try:
                undo_record.EndCustomRecord()
            except Exception:
                pass
