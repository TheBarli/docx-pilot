# Implementation Plan: Document Refinement & Typographic Sanitization

**Target Document**: [documents/Tugas 4 (new).docx](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/documents/Tugas%204%20%28new%29.docx)  
**Baseline Backup**: [documents/Tugas 4 (new)_backup.docx](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/documents/Tugas%204%20%28new%29_backup.docx)

---

## 1. Diagnostic Findings & Target Scope

Deep XML analysis of the current document revealed 4 distinct areas requiring refinement:

| Dimension | Current State | Target Specification |
| :--- | :--- | :--- |
| **Abnormal Bold Text** | **57 long body paragraphs (>100 chars)** are 100% bolded (e.g., full explanation blocks, simulation procedures, conclusions) | Strip full-paragraph bolding. Retain bold **only** on headings, label prefixes (`Jawaban:`, `Kesimpulan:`, `Analisis:`), and table headers. |
| **Empty Tables** | **Table 9** is a 2×2 grid with 0 text and 0 images | **Delete Table 9** completely from document body. |
| **Image Container Tables** | **Table 1** and **Table 3** contain 0 text and exist solely to wrap a `<w:drawing>` | **Extract the images** into clean centered paragraphs and **delete the container tables**. |
| **Multiple Choice Alignment** | Stretched out across lines due to justified alignment (`jc="both"`) | **Enforce clean left-alignment (`jc="left"`)** for all multiple-choice questions and options. |
| **Font Size Uniformity** | Fragmented inline overrides: `7.0 pt` (41), `9.0 pt` (30), `13.0 pt` (221), `13.5 pt` (56), `17.0 pt` (8) | **Strict semantic hierarchy**: Title (14–16pt), H1 (14pt), H2 (12pt), Body (12pt), Tables (11pt), Captions (10–11pt). |

---

## 2. Detailed Technical Execution Plan

### Pillar 1: Surgical Un-bolding of Normal Body Texts
1. **Classification Rule**:
   - A paragraph is classified as a **body paragraph** if it is not `Heading1`, `Heading2`, and is not on the cover page.
2. **Selective Bold Retention Algorithm**:
   - When a body paragraph contains a standard prompt or lead-in prefix (e.g., `Jawaban:`, `Kesimpulan:`, `Analisis:`, `Cara membuat di Packet Tracer:`, `Perangkat:`, `Gambar X – ...:`):
     - Keep the **prefix label in bold**.
     - Strip `<w:b>` and `<w:bCs>` from the subsequent narrative explanation runs.
   - For all other body paragraphs with text length > 80 characters:
     - Remove `<w:b>` and `<w:bCs>` elements from runs, restoring regular font weight.
3. **Preserve Legitimate Bold**:
   - Retain bold on `Heading 1` (14pt bold) and `Heading 2` (12pt bold).
   - Retain bold on table header cells (row 0).
   - Retain bold on short standalone question labels (e.g., `1. Soal:`, `Kasus 1:`).

---

### Pillar 2: Table Sanitization & Image Extraction
1. **Purge Empty Tables**:
   - Detect tables with text length == 0 and image count == 0.
   - **Target**: `Table 9` (2 rows, 4 cells, completely blank).
   - Action: Remove the `<w:tbl>` node from the DOM.
2. **Extract Images from Container Tables**:
   - Detect tables where text length < 15 and image count > 0.
   - **Targets**: 
     - `Table 1` (2×2 grid containing 1 drawing in row 1, cell 1; 0 text).
     - `Table 3` (4×2 grid containing 1 drawing in row 0, cell 1; 0 text).
   - Extraction Procedure:
     1. Locate the `<w:drawing>` element inside the table cell.
     2. Create a clean standalone paragraph:
        ```xml
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="160" w:after="160"/>
          </w:pPr>
          <w:r>
            <!-- Extracted <w:drawing> inserted here -->
          </w:r>
        </w:p>
        ```
     3. Insert the new image paragraph immediately prior to the table in `w:body`.
     4. Delete the `<w:tbl>` container node.
3. **Result**:
   - Document tables reduce from 22 to **19 true data tables**, with images integrated naturally into the document layout.

---

### Pillar 3: Exemption & Left-Alignment for Multiple Choice Questions
1. **Detection Logic**:
   - Identify paragraphs matching multiple-choice patterns:
     - Regex `^[A-E]\.\s+` (individual option lines like `A. Star`, `B. Ring`).
     - Inline concatenated options (e.g., `A. ... B. ... C. ...`).
     - Matching question prompt tables / items.
2. **Alignment Rule**:
   - Explicitly set `<w:jc w:val="left"/>` on `<w:pPr>` for all detected multiple-choice paragraphs.
   - Prevent Word from forcing justified alignment, eliminating unnatural inter-word stretching.

---

### Pillar 4: Semantic Font Size Uniformity
Standardize font size attributes (`<w:sz>` and `<w:szCs>`) strictly by text type, stripping arbitrary inline point sizes:

| Element Type | Font Face | Size in Points | Size in Twips (`w:sz`) | Weight / Style | Alignment |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Main Document Title** | Times New Roman | **16 pt** | `sz="32"` | Bold | Centered |
| **Subtitle / Course Info** | Times New Roman | **12–14 pt** | `sz="24"`–`sz="28"` | Bold | Centered |
| **Heading 1 (Sections A–W)** | Times New Roman | **14 pt** | `sz="28"` | Bold, Black | Left |
| **Heading 2 (Subsections)** | Times New Roman | **12 pt** | `sz="24"` | Bold, Black | Left |
| **Normal Body Paragraphs** | Times New Roman | **12 pt** | `sz="24"` | Regular | Justified |
| **Multiple Choice Items** | Times New Roman | **12 pt** | `sz="24"` | Regular | Left |
| **Table Header Row** | Times New Roman | **11 pt** | `sz="22"` | Bold | Left / Center |
| **Table Data Cells** | Times New Roman | **11 pt** | `sz="22"` | Regular | Left / Center |
| **Figure Captions (`Gambar X`)** | Times New Roman | **11 pt** | `sz="22"` | Italic / Bold prefix | Centered |

*Action*: Strip all orphan `<w:sz w:val="14"/>` (7pt), `<w:sz w:val="18"/>` (9pt), `<w:sz w:val="26"/>` (13pt), `<w:sz w:val="27"/>` (13.5pt), and `<w:sz w:val="34"/>` (17pt) from body runs so they inherit the clean 12pt standard.

---

## 3. Automated Validation Protocol

Upon script completion, verify:
1. **Table Count**: Exactly 19 tables remain (1 empty table deleted, 2 image containers replaced with centered image paragraphs).
2. **Abnormal Bold Count**: 0 full-paragraph bolding in paragraphs > 100 characters.
3. **Multiple Choice Alignment**: All option lines verified with `jc="left"`.
4. **Font Size Distribution**: All body runs normalized to `12.0 pt` (or style-inherited), all table runs normalized to `11.0 pt`.
5. **Schema & Integrity**: Document loads in `python-docx` and MS Word without error.

---

## 4. Opinion & Recommendations for a Truly Uniform, Clean Document

Here is my expert perspective on how to achieve maximum typographic elegance for this document:

### 1. The Power of "Lead-In Bold"
A common mistake in student reports is making entire answers bold to make them stand out. This fatigues the reader's eye and makes the page look heavy.
* **Recommendation**: Use **lead-in bolding**. For example:
  * *Bad (Current)*: **Jawaban: Redundancy sudah melekat secara alami pada topologi mesh karena setiap perangkat memiliki banyak jalur alternatif sehingga komunikasi tetap berjalan.**
  * *Good (Clean)*: **Jawaban:** Redundancy sudah melekat secara alami pada topologi mesh karena setiap perangkat memiliki banyak jalur alternatif sehingga komunikasi tetap berjalan.
  * This creates immediate visual hierarchy: the eye quickly scans to the label, while the explanation reads comfortably in regular weight.

### 2. Centering Extracted Images with Dedicated Captions
Tables should only ever be used for structured multi-column tabular data, never as makeshift containers for screenshots.
* **Recommendation**: When extracting the images from `Table 1` and `Table 3`:
  * Place the `<w:drawing>` inside a centered paragraph (`jc="center"`).
  * Immediately follow it with a clean caption paragraph (e.g., *Gambar 1 – Topologi Star pada Cisco Packet Tracer*), set to 10.5–11pt italic, centered.
  * Inject `<w:keepNext/>` on the image paragraph so the image and its caption are never separated across a page break.

### 3. Left-Aligning Options Prevents "Justification Tears"
In Microsoft Word, short lines that contain tab stops or multiple options (e.g. `A. Star    B. Ring    C. Mesh`) look disastrous when justified. Word forces massive 2-inch spaces between words to make the text reach the right margin.
* **Recommendation**: Exempting multiple choice options from justification and locking them to `jc="left"` is 100% correct typography and will instantly make the test sections look professional.

### 4. Consistent 12pt / 11pt Rhythm
Eliminating the random 13pt and 13.5pt artifacts (likely leftover from Google Docs / WPS font scaling during copy-paste) will restore a unified reading rhythm. In formal academic formatting, having 12pt for text and 11pt for tables is the gold standard.
