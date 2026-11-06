"""Paragraph index mapping, visibility filtering, and outline indexing for Word COM documents."""

from typing import Any, Dict, List

from src.core.exceptions import ERROR_MESSAGES, LiveWordError

WD_WITHIN_TABLE = 12
WD_REVISION_DELETE = 2
WD_OUTLINE_LEVEL_BODY = 10


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

