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

    visible_indices = [i for i in range(1, total_p + 1) if i not in table_p_set]
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
                raise NotImplementedError("Table inspection in live mode is scheduled for Phase 2.")

            if mode == "outline":
                lines = [
                    f"# Outline: {path.name}",
                    f"- **Total Paragraphs**: {num_paragraphs} (indices: [p:0] to [p:{max(0, num_paragraphs - 1)}])",
                    f"- **Total Tables**: {num_tables}",
                    "",
                    "### Structural Elements (0-based indices):",
                ]
                try:
                    items = doc.GetCrossReferenceItems(1)  # wdRefTypeHeading = 1
                except Exception:
                    items = ()

                if items:
                    com_to_zero = {com_i: idx for idx, com_i in enumerate(body_indices)}
                    rng = doc.Content
                    f = rng.Find
                    for heading_text in items:
                        clean = heading_text.strip()
                        if not clean:
                            continue
                        f.ClearFormatting()
                        if f.Execute(FindText=clean, Forward=True):
                            p_obj = rng.Paragraphs.Item(1)
                            com_i = doc.Range(0, p_obj.Range.End).Paragraphs.Count
                            zero_idx = com_to_zero.get(com_i, 0)
                            style_name = str(p_obj.Style.NameLocal)
                            lines.append(f"- `[p:{zero_idx}]` **({style_name})**: \"{clean[:80]}\"")
                            rng.Collapse(0)
                else:
                    for zero_idx, com_i in enumerate(body_indices):
                        p = doc.Paragraphs.Item(com_i)
                        style_name = str(p.Style.NameLocal)
                        outline_lvl = getattr(p, "OutlineLevel", WD_OUTLINE_LEVEL_BODY)
                        if outline_lvl < WD_OUTLINE_LEVEL_BODY or "heading" in style_name.lower():
                            txt = p.Range.Text.rstrip("\r\n\x07")
                            lines.append(f"- `[p:{zero_idx}]` **({style_name})**: \"{txt[:80]}\"")

                for t_idx in range(1, num_tables + 1):
                    tbl = tables.Item(t_idx)
                    lines.append(f"- `[table:{t_idx - 1}]` **Table ({tbl.Rows.Count} rows x {tbl.Columns.Count} cols)**")

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
                body_rng = doc.Content
                find_obj = body_rng.Find
                find_obj.ClearFormatting()
                find_obj.Text = find_text
                find_obj.MatchCase = match_case
                find_obj.Forward = True
                find_obj.Wrap = 0  # wdFindStop

                # Locate each match and replace literally
                while find_obj.Execute():
                    # Skip if inside tracked deletion
                    in_del = False
                    try:
                        for rev in body_rng.Revisions:
                            if rev.Type == WD_REVISION_DELETE:
                                in_del = True
                                break
                    except Exception:
                        pass

                    if not in_del:
                        body_rng.Text = replace_text
                        count += 1

                    # Collapse to end and continue search forward
                    body_rng.Collapse(0)  # wdCollapseEnd

            return {
                "status": "success",
                "file_path": str(path),
                "matches_replaced": count,
                "find_text": find_text,
                "replace_text": replace_text,
            }

        return com_worker.run(_fnr)
