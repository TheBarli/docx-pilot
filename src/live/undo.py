"""Undo record and Track Changes context managers for live Word COM mutations."""

from contextlib import contextmanager
from typing import Any, Generator, Optional


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

