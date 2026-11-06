"""Backward-compatibility facade for Live Word COM engine.

Re-exports core exceptions, paths, worker singleton, undo wrappers,
router functions, indexer helpers, and LiveWordEngine from src.core and src.live packages.
"""

from src.core.exceptions import ERROR_MESSAGES, LiveWordError
from src.core.paths import _norm, paths_match
from src.live.engine import LiveWordEngine, _hex_to_bgr
from src.live.indexer import (
    WD_OUTLINE_LEVEL_BODY,
    WD_REVISION_DELETE,
    WD_WITHIN_TABLE,
    _build_visible_map,
    _com_idx,
    _get_local_outline,
    _visible_map,
)
from src.live.router import find_open_document, has_word_lock_file, pick_backend
from src.live.undo import live_call, tracking
from src.live.worker import HAS_COM, ComWorker, com_worker

__all__ = [
    "HAS_COM",
    "LiveWordEngine",
    "ComWorker",
    "com_worker",
    "find_open_document",
    "pick_backend",
    "paths_match",
    "has_word_lock_file",
    "LiveWordError",
    "ERROR_MESSAGES",
    "_norm",
    "tracking",
    "live_call",
    "_build_visible_map",
    "_visible_map",
    "_com_idx",
    "_get_local_outline",
    "_hex_to_bgr",
    "WD_WITHIN_TABLE",
    "WD_REVISION_DELETE",
    "WD_OUTLINE_LEVEL_BODY",
]
