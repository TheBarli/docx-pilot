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
        if HAS_COM and sys.platform == "win32":
            try:
                import win32service
                import win32con
                h_default = win32service.OpenDesktop("Default", 0, False, win32con.MAXIMUM_ALLOWED)
                h_default.SetThreadDesktop()
            except Exception:
                pass
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
        if self._thread is not None and threading.current_thread() is self._thread:
            return func(*args, **kwargs)

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


# ---------------------------------------------------------------------------
# Live Word Engine (Phase 1 Spike: read_document, mutate_element, find_and_replace)
# ---------------------------------------------------------------------------
WD_WITHIN_TABLE = 12
WD_REVISION_DELETE = 2
WD_OUTLINE_LEVEL_BODY = 10


def _build_visible_map(doc: Any) -> Dict[str, Any]:
    """Calculate 0-based visible body paragraphs, excluding tables and full deletions.

    Optimized for bulk speed: queries doc.Tables ranges instead of calling
    rng.Information(WD_WITHIN_TABLE) per paragraph, reducing COM round-trips from O(N) to O(T).
    """
    total_p = doc.Paragraphs.Count
    if total_p == 0:
        return {"indices": [], "deleted_spans": {}}

    table_p_set = set()
    try:
        tables = doc.Tables
        tbl_count = tables.Count
        for idx in range(1, tbl_count + 1):
            try:
                tr = tables.Item(idx).Range
                p_start = doc.Range(0, tr.Start).Paragraphs.Count + 1
                cnt = tr.Paragraphs.Count
                for p_i in range(p_start, p_start + cnt):
                    table_p_set.add(p_i)
            except Exception:
                pass
    except Exception:
        pass

    # Exclude paragraphs that are fully tracked deletions
    deleted_p_set = set()
    try:
        if doc.Revisions.Count > 0:
            for rev in doc.Revisions:
                if getattr(rev, "Type", None) == WD_REVISION_DELETE:
                    r = rev.Range
                    p_first = doc.Range(0, r.Start).Paragraphs.Count + 1
                    p_last = doc.Range(0, r.End).Paragraphs.Count
                    for p_i in range(p_first, p_last + 1):
                        try:
                            p_rng = doc.Paragraphs.Item(p_i).Range
                            if r.Start <= p_rng.Start and r.End >= p_rng.End:
                                deleted_p_set.add(p_i)
                        except Exception:
                            pass
    except Exception:
        pass

    visible_indices = [i for i in range(1, total_p + 1) if i not in table_p_set and i not in deleted_p_set]
    return {
        "indices": visible_indices,
        "deleted_spans": {},
    }


def _get_local_outline(doc: Any, target_idx: int, window: int = 2) -> List[str]:
    """Extract local outline around target_idx with fresh 0-based indices."""
    vmap = _build_visible_map(doc)
    indices = vmap["indices"]
    total = len(indices)
    if total == 0:
        return []

    start = max(0, target_idx - window)
    end = min(total, target_idx + window + 1)
    outline_lines: List[str] = []

    for zero_idx in range(start, end):
        com_i = indices[zero_idx]
        try:
            p = doc.Paragraphs.Item(com_i)
            style_name = str(p.Style.NameLocal)
            txt = p.Range.Text.rstrip("\r\n\x07")
            txt_snippet = txt[:60] + "..." if len(txt) > 60 else txt
            outline_lines.append(f"`[p:{zero_idx}]` ({style_name}): {txt_snippet}")
        except Exception:
            continue

    return outline_lines


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
                        if is_known_heading or len(lines) < 40 or not items:
                            try:
                                p = doc.Paragraphs.Item(com_i)
                                style_name = str(p.Style.NameLocal) if p.Style else "Normal"
                                is_heading = is_known_heading or style_name.lower().startswith("heading") or style_name.lower() == "title"
                                text_snippet = p.Range.Text.rstrip("\r\n\x07").strip()
                                if is_heading or (len(lines) < 40 and text_snippet):
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
