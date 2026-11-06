"""XML element ordering and document loading helpers for headless Word operations."""

from pathlib import Path
from typing import List, Tuple

from docx import Document
from lxml import etree

from src.core.paths import normalize_path


def reorder_children(parent: etree._Element, order_list: List[str]) -> None:
    """Sort children of an XML element according to ECMA-376 sequence specification."""
    order_map = {tag: i for i, tag in enumerate(order_list)}
    children = list(parent)
    children.sort(key=lambda el: order_map.get(el.tag.split("}")[-1] if "}" in el.tag else el.tag, 999))
    for c in list(parent):
        parent.remove(c)
    for c in children:
        parent.append(c)


def get_or_create_document(file_path: str, create_if_missing: bool = False) -> Tuple[Document, Path]:
    """Load existing document or create a new one if permitted.
    
    Returns tuple of (Document, resolved_Path).
    """
    path = normalize_path(file_path)
    if not path.exists():
        if create_if_missing:
            path.parent.mkdir(parents=True, exist_ok=True)
            doc = Document()
            doc.save(str(path))
            return doc, path
        raise FileNotFoundError(f"File not found: {path}")
    return Document(str(path)), path
