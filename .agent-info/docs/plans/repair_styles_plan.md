# Surgical Implementation Plan: Repair & Formatting Alignment for `Tugas 4.docx`

Based on the definitive root cause analysis in [problem-analysis-bloat-fixing.md](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/.agent-info/problem-analysis-bloat-fixing.md), all full-transplant and document-body rewrites are rejected in favor of a **surgical, zero-blast-radius XML patch**.

---

## 1. Executive Summary & Root Cause Synthesis

### A. The "Styles 1" Unreadable Error
* **Root Cause**: 24 legacy table style elements (`a`, `a0` through `af6`) in `word/styles.xml` are completely missing the mandatory `<w:name>` child tag required by ECMA-376 (ISO/IEC 29500).
* **The Fix**: Inject `<w:name w:val="Table {styleId}"/>` as the first child of each malformed `<w:style>` node.
* **Blast Radius**: Zero changes to table contents or paragraphs. No rewriting of table style IDs across `word/document.xml`.

### B. The File Size Bloat (23.3 MB $\to$ ~2.5 MB)
* **Root Cause 1**: Full system font binaries (including 16 MB Korean font `Gungsuh`) dumped into `word/fonts/` because `word/settings.xml` enabled `embedTrueTypeFonts` with `saveSubsetFonts=false`.
* **Root Cause 2**: 39 images in `word/media/` were re-saved by Word using `compress_type: 0` (`ZIP_STORED` raw uncompressed bitmaps) instead of `ZIP_DEFLATED`.
* **The Fix**: 
  1. Strip `<w:embedTrueTypeFonts/>`, `<w:saveSubsetFonts/>`, and `<w:doNotEmbedSystemFonts/>` from `word/settings.xml` (prevents Word re-embedding on future `Ctrl+S`).
  2. Strip `<w:embedRegular>`, `<w:embedBold>`, `<w:embedItalic>` from `word/fontTable.xml`.
  3. Remove `word/fonts/` and `word/_rels/fontTable.xml.rels`.
  4. Repack the entire package using `zipfile.ZIP_DEFLATED`.

### C. Formatting & Layout Alignment (from `EVALUASI MATERI 3`)
* **Margins & Dimensions**:
  * Set `<w:pgMar>` to standard 4-4-3-3 cm:
    * `top="2268"` twips (4.00 cm)
    * `left="2268"` twips (4.00 cm)
    * `bottom="1701"` twips (3.00 cm)
    * `right="1701"` twips (3.00 cm)
  * Set `<w:pgSz>` to Letter size (`w="12240"`, `h="15840"`).
  * **Surgical Target**: Modify all `<w:sectPr>` elements across the document tree (`tree.findall('.//{...}sectPr')`). This guarantees margins are applied across all sections if section breaks exist, rather than only the final section. Leave all `<w:p>` and `<w:tbl>` nodes untouched.
* **Typography Alignment**:
  * Update `<w:style w:styleId="Normal">` in `word/styles.xml` to Times New Roman 12pt (`sz="24"`), 1.5 line spacing (`line="360"`, `after="0"`), justified alignment (`jc="both"`).

---

## 2. Comparison of Architectures

| Evaluation Dimension | Legacy Full Transplant (Rejected) | Flawed "No-Touch" Hybrid (Rejected) | Approved Surgical Architecture |
| :--- | :--- | :--- | :--- |
| **"Styles 1" Error** | Fixed | Fixed | **Fixed (100% ECMA-376 compliant)** |
| **`document.xml` Edits** | Heavy (Rewrites 24 table style IDs) | Claims none (Factually impossible for margins) | **Micro-surgical** (Only touches `<w:sectPr>`, tables & paragraphs untouched) |
| **Table Formatting Risk** | High (Destroys custom borders/padding) | Zero | **Zero (Existing table properties 100% preserved)** |
| **Page Layout (4-4-3-3 cm)** | Applied | Failed | **Applied (2268 / 1701 twips explicitly set)** |
| **Word Re-Save Trap** | Exposed (fonts re-embed on Ctrl+S) | Exposed | **Eliminated (settings & fontTable embed flags stripped)** |
| **Final File Size** | Variable | Variable | **Lean ~2.5 MB (`ZIP_DEFLATED`)** |

---

## 3. Concrete Action Items

1. **Input Baseline**: Work from untouched reference [documents/Tugas 4_backup.docx](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/documents/Tugas%204_backup.docx).
2. **Execute Surgical Script**: Run the verified script [.agent-info/test_surgical_repair.py](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/.agent-info/test_surgical_repair.py) targeting [documents/Tugas 4.docx](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/documents/Tugas%204.docx).
3. **Automated Validation**:
   - Verify `missing w:name == 0`.
   - Verify all 24 tables and paragraphs are intact.
   - Verify page dimensions (Letter, 4-4-3-3 cm).
   - Verify final package size is ~2.5 MB.
4. **User Verification**: User opens `documents/Tugas 4.docx` in Microsoft Word to confirm no recovery/repair dialog appears and file saves without bloat.
