"""Apply formatting and styling from EVALUASI MATERI 3 (1).docx to Tugas 4.docx.

Key standards copied:
1. Margins: Top 4cm (113.4pt), Left 4cm (113.4pt), Bottom 3cm (85.05pt), Right 3cm (85.05pt).
2. Page Size: Letter (21.59 cm x 27.94 cm / 612pt x 792pt).
3. Font: Times New Roman throughout.
4. Line Spacing: 1.5 line spacing (w:line=360).
5. Paragraph Spacing: Space before = 0pt, Space after = 0pt.
6. Body Alignment: Justified (both), except Center for cover page / titles.
7. Cover Page: Centered title, lecturer, group name, Untirta logo, prodi/fakultas/universitas/year.
8. Tables: Times New Roman, bold headers, clean grid borders.
"""

from pathlib import Path
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC_FILE = Path("documents/EVALUASI MATERI 3 (1).docx")
DST_FILE = Path("documents/Tugas 4.docx")


def apply_styling():
    print(f"Loading {DST_FILE}...")
    doc = docx.Document(str(DST_FILE))

    # 1. Page Setup & Margins (Indonesian Academic Standard 4-4-3-3 cm)
    print("Applying section margins and page size (Letter, Top 4cm, Left 4cm, Bottom 3cm, Right 3cm)...")
    for sec in doc.sections:
        sec.top_margin = Cm(4.0)
        sec.left_margin = Cm(4.0)
        sec.bottom_margin = Cm(3.0)
        sec.right_margin = Cm(3.0)
        sec.page_width = Pt(612.0)   # Letter 8.5 in
        sec.page_height = Pt(792.0)  # Letter 11 in
        sec.header_distance = Cm(1.5)
        sec.footer_distance = Cm(1.5)

    # 2. Update Document Normal Style & Defaults
    print("Configuring Normal style (Times New Roman, 12pt, 1.5 line spacing, 0pt after/before)...")
    try:
        style_normal = doc.styles["Normal"]
        style_normal.font.name = "Times New Roman"
        style_normal.font.size = Pt(12)
        style_normal.paragraph_format.line_spacing = 1.5
        style_normal.paragraph_format.space_before = Pt(0)
        style_normal.paragraph_format.space_after = Pt(0)
        style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    except Exception as e:
        print(f"Notice when setting Normal style: {e}")

    # Set styles for Headings if present
    for h_name in ["Heading 1", "Heading 2", "Heading 3"]:
        try:
            h_style = doc.styles[h_name]
            h_style.font.name = "Times New Roman"
            h_style.font.size = Pt(12)
            h_style.font.bold = True
            h_style.font.color.rgb = RGBColor(0, 0, 0)
            h_style.paragraph_format.line_spacing = 1.5
            h_style.paragraph_format.space_before = Pt(6)
            h_style.paragraph_format.space_after = Pt(0)
        except Exception:
            pass

    # 3. Process Paragraphs
    print(f"Formatting {len(doc.paragraphs)} paragraphs...")
    
    # Identify cover page boundary: look for year / page break or paragraph count
    cover_page_end = 20
    for idx, p in enumerate(doc.paragraphs[:35]):
        t = p.text.strip().upper()
        if "SEPTEMBER 2026" in t or "2026" in t and ("UNIVERSITAS" in doc.paragraphs[max(0, idx-1)].text.upper() or "TEKNIK" in doc.paragraphs[max(0, idx-1)].text.upper()):
            cover_page_end = idx + 1
            break

    print(f"Cover page detected up to paragraph index {cover_page_end}")

    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        is_cover = idx < cover_page_end

        # Ensure all runs use Times New Roman
        for r in p.runs:
            r.font.name = "Times New Roman"
            # Normalize font size to 12pt unless explicit sub/superscript
            if not r.font.size or r.font.size.pt not in (8, 9, 10, 11):
                r.font.size = Pt(12)
            # Default text color black
            if not r.font.color.rgb:
                r.font.color.rgb = RGBColor(0, 0, 0)

        # Cover Page Formatting
        if is_cover:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            # Use 1.0 or 1.15 line spacing for member list and details on cover page so it stays on 1 page
            p.paragraph_format.line_spacing = 1.15
            continue

        # Body Paragraphs Formatting
        # Check if paragraph is a heading or title
        is_heading_text = False
        upper_text = text.upper()
        if (
            upper_text.startswith("EVALUASI MATERI")
            or upper_text in ("FREE TEST", "POST TEST", "PERTANYAAN URAIAN", "KOGNITIF", "AFEKTIF", "PSIKOMETRIK", "SOAL MENCOCOKAN")
            or (len(text) < 60 and any(r.bold for r in p.runs) and (text.startswith("1.") or text.startswith("A.") or text.startswith("B.") or text.startswith("BAB ")))
        ):
            is_heading_text = True

        if is_heading_text:
            if "EVALUASI MATERI" in upper_text or "FREE TEST" in upper_text:
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(0)
            for r in p.runs:
                r.font.bold = True
        else:
            # Standard Body Paragraph
            # If paragraph was explicitly centered in original (e.g. figure caption or equation)
            if p.alignment == WD_ALIGN_PARAGRAPH.CENTER or (len(text) < 50 and text.lower().startswith("gambar")):
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Standard 1.5 line spacing, 0pt before/after
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

    # 4. Process Tables
    print(f"Formatting {len(doc.tables)} tables...")
    for t_idx, tbl in enumerate(doc.tables):
        try:
            tbl.style = "Table Grid"
        except Exception:
            pass

        for row_idx, row in enumerate(tbl.rows):
            is_header = (row_idx == 0)
            for cell in row.cells:
                for cp in cell.paragraphs:
                    # Clean spacing inside table cells
                    cp.paragraph_format.space_before = Pt(0)
                    cp.paragraph_format.space_after = Pt(0)
                    cp.paragraph_format.line_spacing = 1.15
                    if is_header:
                        cp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for cr in cp.runs:
                        cr.font.name = "Times New Roman"
                        cr.font.size = Pt(11)  # 11pt for tables matching academic standard
                        if is_header:
                            cr.font.bold = True
                        if not cr.font.color.rgb:
                            cr.font.color.rgb = RGBColor(0, 0, 0)

    print(f"Saving updated document to {DST_FILE}...")
    doc.save(str(DST_FILE))
    print("SUCCESS: Styling and formatting applied successfully!")


if __name__ == "__main__":
    apply_styling()
