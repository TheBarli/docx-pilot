"""Live Microsoft Word backend engine via COM interface."""

from __future__ import annotations

from src.live.tools_edit import EditTools
from src.live.tools_format import FormatTools, _hex_to_bgr
from src.live.tools_read import ReadTools

__all__ = ["LiveWordEngine", "_hex_to_bgr"]


class LiveWordEngine(ReadTools, EditTools, FormatTools):
    """Live Word backend operations over Microsoft Word COM interface."""

    pass
