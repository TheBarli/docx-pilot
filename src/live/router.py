"""Live detection, open document search, and backend routing between live and headless."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any, Optional, Tuple, Union

from src.core.exceptions import ERROR_MESSAGES, LiveWordError
from src.core.paths import paths_match
from src.live.worker import HAS_COM, com_worker

try:
    if sys.platform == "win32":
        import win32com.client
except ImportError:
    pass


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

