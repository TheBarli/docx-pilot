"""Quick Demo Script for Word DOCX FastMCP Server.

Run this script to test the server tools end-to-end and generate a real Word document:
    .venv\Scripts\python.exe demo.py
"""

from pathlib import Path
from src.oxml_engine import DocxEngine

DOCS_DIR = Path(__file__).parent / "documents"
OUTPUT_DOC = DOCS_DIR / "demo_document.docx"

print("=" * 60)
print("1. Creating document with compose_document (batch write)...")
print("=" * 60)

actions = [
    {"type": "heading", "text": "Annual AI Integration Report", "level": 1},
    {
        "type": "paragraph",
        "text": "This document was generated and manipulated directly via FastMCP and python-docx.",
        "bold": True,
    },
    {"type": "list_item", "text": "Direct XML-level element mutation", "style": "List Bullet"},
    {"type": "list_item", "text": "0-based indexing for zero hallucinations", "style": "List Bullet"},
    {"type": "list_item", "text": "Consolidated 8-tool MCP architecture", "style": "List Bullet"},
    {
        "type": "table",
        "headers": ["Quarter", "Status", "Completion"],
        "rows": [
            ["Q1", "Completed", "100%"],
            ["Q2", "In Progress", "75%"],
            ["Q3", "Planned", "0%"],
        ],
    },
    {"type": "heading", "text": "Summary & Next Steps", "level": 2},
    {"type": "paragraph", "text": "Review the metrics above before final export."},
]

res = DocxEngine.compose_document(str(OUTPUT_DOC), actions, create_if_missing=True)
print(f"-> Created: {res['file_path']}")
print(f"-> Actions executed: {res['actions_executed']}")
print(f"-> Paragraphs: {res['total_paragraphs']}, Tables: {res['total_tables']}")

print("\n" + "=" * 60)
print("2. Reading Document Outline (0-based indexing):")
print("=" * 60)
outline = DocxEngine.read_document(str(OUTPUT_DOC), mode="outline")
print(outline)

print("\n" + "=" * 60)
print("3. Performing Low-Level XML Mutation: Insert Paragraph before [p:1]...")
print("=" * 60)
mutate_res = DocxEngine.mutate_element(
    str(OUTPUT_DOC),
    target_index=1,
    action="insert_before",
    content={"type": "paragraph", "text": ">> [INSERTED VIA OXML MUTATION] Executive Summary <<", "bold": True},
)
print(f"-> Mutation status: {mutate_res['status']}, New paragraph count: {mutate_res['total_paragraphs']}")

print("\n" + "=" * 60)
print("4. Performing Find & Replace ('Planned' -> 'Fast-Tracked')...")
print("=" * 60)
replace_res = DocxEngine.find_and_replace(
    str(OUTPUT_DOC),
    find_text="Planned",
    replace_text="Fast-Tracked",
    scope="all",
)
print(f"-> Replacements made: {replace_res['matches_replaced']} at {replace_res['sample_locations']}")

print("\n" + "=" * 60)
print("5. Reading Updated Document Full Dump:")
print("=" * 60)
full_dump = DocxEngine.read_document(str(OUTPUT_DOC), mode="full")
print(full_dump)

print("\n" + "=" * 60)
print(f"SUCCESS! You can open 'documents/{OUTPUT_DOC.name}' in Microsoft Word to see the final file.")
print("=" * 60)
