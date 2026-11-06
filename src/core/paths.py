"""Path normalization and matching utilities for Word document paths."""

import os
import sys
from pathlib import Path
from typing import Union

from src.core.constants import PROJECT_ROOT

HAS_COM = False
try:
    if sys.platform == "win32":
        import win32api
        HAS_COM = True
except ImportError:
    HAS_COM = False


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


def normalize_path(file_path: str) -> Path:
    """Normalize file path and resolve to absolute path, checking workspace root if relative."""
    p = Path(file_path).expanduser()
    if p.is_absolute() and p.exists():
        return p.resolve()
    cwd_candidate = (Path.cwd() / p).resolve()
    if cwd_candidate.exists():
        return cwd_candidate
    root_candidate = (PROJECT_ROOT / p).resolve()
    if root_candidate.exists() or not p.is_absolute():
        return root_candidate
    return p.resolve()
