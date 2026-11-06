"""Backward-compatibility facade for OpenXML headless engine.

Re-exports core constants, path utilities, XML helpers, and DocxEngine
from src.core and src.headless packages.
"""

from src.core.constants import (
    OUTLINE_LINE_LIMIT,
    PPR_ORDER,
    PROJECT_ROOT,
    RPR_ORDER,
    STYLE_ORDER,
    W_NS,
)
from src.core.paths import normalize_path
from src.headless.engine import DocxEngine
from src.headless.xml_helpers import get_or_create_document, reorder_children

__all__ = [
    "DocxEngine",
    "normalize_path",
    "get_or_create_document",
    "reorder_children",
    "PROJECT_ROOT",
    "W_NS",
    "PPR_ORDER",
    "RPR_ORDER",
    "STYLE_ORDER",
    "OUTLINE_LINE_LIMIT",
]
