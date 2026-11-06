"""Core OpenXML and python-docx Engine for headless operations assembled via mixins."""

from __future__ import annotations

from src.headless.tools_edit import EditTools
from src.headless.tools_format import FormatTools
from src.headless.tools_read import ReadTools


class DocxEngine(ReadTools, EditTools, FormatTools):
    """Low-level engine for Word document inspection, composition, and XML mutation."""
    pass
