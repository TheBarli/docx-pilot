"""Unit tests for oxml_engine.py."""

from pathlib import Path
import pytest
from src.oxml_engine import DocxEngine, normalize_path


@pytest.fixture
def temp_docx(tmp_path: Path) -> str:
    """Fixture providing temporary docx path."""
    return str(tmp_path / "test_doc.docx")


def test_compose_and_read(temp_docx: str):
    """Test composing a document with batch actions and reading all modes."""
    actions = [
        {"type": "heading", "text": "Document Title", "level": 1},
        {"type": "paragraph", "text": "This is introductory paragraph.", "bold": True},
        {"type": "list_item", "text": "First bullet point", "style": "List Bullet"},
        {"type": "list_item", "text": "Second bullet point", "style": "List Bullet"},
        {
            "type": "table",
            "headers": ["Item", "Quantity", "Price"],
            "rows": [["Widget A", "10", "$5.00"], ["Widget B", "25", "$12.50"]],
        },
        {"type": "page_break"},
        {"type": "heading", "text": "Appendix", "level": 2},
        {"type": "paragraph", "text": "Final notes."},
    ]

    res = DocxEngine.compose_document(temp_docx, actions, create_if_missing=True)
    assert res["status"] == "success"
    assert res["actions_executed"] == 8
    assert res["total_tables"] == 1

    # 1. Test Outline mode (verifying 0-based [p:0], [table:0])
    outline = DocxEngine.read_document(temp_docx, mode="outline")
    assert isinstance(outline, str)
    assert "[p:0]" in outline
    assert "[table:0]" in outline
    assert "Document Title" in outline

    # 2. Test Full mode
    full = DocxEngine.read_document(temp_docx, mode="full")
    assert isinstance(full, str)
    assert "`[p:0]` (Heading 1): Document Title" in full
    assert "`[table:0]` Table (3x3):" in full

    # 3. Test Table mode (0-based table_index=0)
    tbl_data = DocxEngine.read_document(temp_docx, mode="table", table_index=0)
    assert isinstance(tbl_data, dict)
    assert tbl_data["table_index"] == 0
    assert tbl_data["rows_count"] == 3
    assert tbl_data["columns_count"] == 3
    assert tbl_data["matrix"][0] == ["Item", "Quantity", "Price"]
    assert tbl_data["matrix"][1] == ["Widget A", "10", "$5.00"]

    # 4. Test Metadata mode
    meta = DocxEngine.read_document(temp_docx, mode="metadata")
    assert isinstance(meta, dict)
    assert meta["total_tables"] == 1


def test_xml_paragraph_mutations(temp_docx: str):
    """Test XML-level insert_before, insert_after, delete, and replace."""
    # Start with 3 paragraphs: P0, P1, P2
    initial_actions = [
        {"type": "paragraph", "text": "Original P0"},
        {"type": "paragraph", "text": "Original P1"},
        {"type": "paragraph", "text": "Original P2"},
    ]
    DocxEngine.compose_document(temp_docx, initial_actions, create_if_missing=True)

    # 1. Insert before [p:1] -> should become new paragraph at index 1
    res1 = DocxEngine.mutate_element(
        temp_docx,
        target_index=1,
        action="insert_before",
        content={"type": "paragraph", "text": "Inserted Before P1"},
    )
    assert res1["status"] == "success"
    assert res1["total_paragraphs"] == 4

    full1 = DocxEngine.read_document(temp_docx, mode="full")
    assert "`[p:0]` (Normal): Original P0" in full1
    assert "`[p:1]` (Normal): Inserted Before P1" in full1
    assert "`[p:2]` (Normal): Original P1" in full1
    assert "`[p:3]` (Normal): Original P2" in full1

    # 2. Insert after [p:0] -> should become new paragraph at index 1
    res2 = DocxEngine.mutate_element(
        temp_docx,
        target_index=0,
        action="insert_after",
        content={"type": "paragraph", "text": "Inserted After P0"},
    )
    assert res2["status"] == "success"
    assert res2["total_paragraphs"] == 5

    full2 = DocxEngine.read_document(temp_docx, mode="full")
    assert "`[p:0]` (Normal): Original P0" in full2
    assert "`[p:1]` (Normal): Inserted After P0" in full2

    # 3. Replace [p:1] in-place
    res3 = DocxEngine.mutate_element(
        temp_docx,
        target_index=1,
        action="replace",
        content={"text": "Replaced Content P1", "bold": True},
    )
    assert res3["status"] == "success"
    full3 = DocxEngine.read_document(temp_docx, mode="full")
    assert "`[p:1]` (Normal): Replaced Content P1" in full3

    # 4. Delete [p:1]
    res4 = DocxEngine.mutate_element(
        temp_docx,
        target_index=1,
        action="delete",
    )
    assert res4["status"] == "success"
    assert res4["remaining_paragraphs"] == 4

    full4 = DocxEngine.read_document(temp_docx, mode="full")
    assert "Replaced Content P1" not in full4


def test_find_and_replace(temp_docx: str):
    """Test scoped find and replace in body and tables."""
    actions = [
        {"type": "paragraph", "text": "The quick brown fox jumps over the lazy dog."},
        {
            "type": "table",
            "headers": ["Animal", "Trait"],
            "rows": [["fox", "quick"], ["dog", "lazy"]],
        },
    ]
    DocxEngine.compose_document(temp_docx, actions, create_if_missing=True)

    # Replace "fox" with "cat"
    res = DocxEngine.find_and_replace(
        temp_docx,
        find_text="fox",
        replace_text="cat",
        match_case=True,
        scope="all",
    )
    assert res["status"] == "success"
    assert res["matches_replaced"] == 2

    full = DocxEngine.read_document(temp_docx, mode="full")
    assert "The quick brown cat jumps over the lazy dog." in full

    tbl = DocxEngine.read_document(temp_docx, mode="table", table_index=0)
    assert tbl["matrix"][1][0] == "cat"
