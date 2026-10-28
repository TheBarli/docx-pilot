"""Debloat documents/Tugas 4.docx by stripping embedded font bloat and recompressing media.

Matches the configuration of EVALUASI MATERI 3 (1).docx:
- Removes w:embedTrueTypeFonts from settings.xml so Word won't re-embed fonts on save.
- Strips 24MB of unneeded system font binaries (Gungsuh, Calibri, Cambria, etc.).
- Compresses media files using ZIP_DEFLATED.
"""

from pathlib import Path
import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import docx

TARGET_FILE = Path("documents/Tugas 4.docx")
BACKUP_FILE = Path("documents/Tugas 4_backup.docx")


def debloat_docx(docx_path: Path):
    print(f"Reading {docx_path}...")
    size_before = os.path.getsize(docx_path)

    # Use a temporary file for atomic write
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(
            tmp_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                # 1. Skip embedded font files completely (matches EVALUASI MATERI 3)
                if item.filename.startswith("word/fonts/") or item.filename == "word/_rels/fontTable.xml.rels":
                    continue

                content = zin.read(item.filename)

                # 2. In settings.xml, remove font embedding tags
                if item.filename == "word/settings.xml":
                    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                    ET.register_namespace("w", ns["w"])
                    root = ET.fromstring(content)
                    for tag in ["w:embedTrueTypeFonts", "w:doNotEmbedSystemFonts", "w:saveSubsetFonts"]:
                        el = root.find(tag, ns)
                        if el is not None:
                            root.remove(el)
                    content = ET.tostring(root, xml_declaration=True, encoding="utf-8")

                # 3. In fontTable.xml, remove font embedding references
                if item.filename == "word/fontTable.xml":
                    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                    ET.register_namespace("w", ns["w"])
                    root = ET.fromstring(content)
                    for font in root.findall("w:font", ns):
                        for embed_tag in ["w:embedRegular", "w:embedBold", "w:embedItalic", "w:embedBoldItalic"]:
                            em = font.find(embed_tag, ns)
                            if em is not None:
                                font.remove(em)
                    content = ET.tostring(root, xml_declaration=True, encoding="utf-8")

                # 4. Enforce ZIP_DEFLATED on all entries (including images)
                new_item = zipfile.ZipInfo(item.filename)
                new_item.compress_type = zipfile.ZIP_DEFLATED
                # Preserve permissions
                new_item.external_attr = item.external_attr
                zout.writestr(new_item, content)

        # Verify integrity by opening with python-docx before replacing target
        test_doc = docx.Document(str(tmp_path))
        p_count = len(test_doc.paragraphs)
        t_count = len(test_doc.tables)
        print(f"Integrity check passed: {p_count} paragraphs, {t_count} tables.")

        # Replace target with debloated file
        shutil.move(str(tmp_path), str(docx_path))
        size_after = os.path.getsize(docx_path)
        saved = size_before - size_after

        print("=== Debloating Results ===")
        print(f"Size Before: {size_before:,} bytes ({size_before / 1024 / 1024:.2f} MB)")
        print(f"Size After:  {size_after:,} bytes ({size_after / 1024 / 1024:.2f} MB)")
        print(f"Space Saved: {saved:,} bytes ({saved / 1024 / 1024:.2f} MB saved - {saved / size_before * 100:.1f}%)")

    finally:
        if tmp_path.exists():
            tmp_path.unlink()


if __name__ == "__main__":
    debloat_docx(TARGET_FILE)
