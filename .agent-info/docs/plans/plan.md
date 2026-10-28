# Comprehensive Plan: Formatting & Sanitization of `Tugas 4 (new).docx`

This plan establishes the complete architectural specification to transform [documents/Tugas 4 (new).docx](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/documents/Tugas%204%20%28new%29.docx) into a clean, uniform, professional academic document matching the standards of [documents/EVALUASI MATERI 3 (1).docx](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/documents/EVALUASI%20MATERI%203%20%281%29.docx).

---

## 1. Document Diagnostic Baseline

Deep XML analysis of `Tugas 4 (new).docx` revealed significant formatting inconsistencies:

| Inspection Dimension | Current State in `Tugas 4 (new).docx` | Target Standard (Evaluasi Aligned) |
| :--- | :--- | :--- |
| **Total Paragraphs** | 1,440 paragraphs | ~1,210 paragraphs (after purging bloat) |
| **Useless Blank Spaces** | **187 empty** + **42 whitespace-only** paragraphs | **0** in document body (spacing handled via `w:spacing`) |
| **Page Breaks** | **13 page breaks** (3 broken mid-sentence/mid-answer) | **1 clean page break** (immediately following cover page) |
| **Font Colors** | Fragmented: `1a1a1a`, `2b2b2b`, `13599c` (blue), `0a3d66` (navy), `555555`, `ffffff` (white) | **100% Pure Black Ink** (`#000000` / `auto`) |
| **Tables** | 22 tables with 22 disparate vendor styles (`Table1`–`Table22`), conflicting borders (`nil` vs `single`), dark blue shading | **Uniform Standard Grid** (Single 0.5pt black border, neutral padding, consistent fonts) |
| **Headings** | Zero formal heading styles applied (raw bold body text) | **Automated Heading Hierarchy** (`Heading 1`, `Heading 2`, `Heading 3`) with `keepNext` |
| **Page Layout & Margins** | A4 (11909 × 16834), 2.54 cm margins | **Letter** (12240 × 15840), **4-4-3-3 cm** (Top 4cm, Left 4cm, Bottom 3cm, Right 3cm) |
| **Font Bloat & Styles** | 22 malformed styles missing `<w:name>`, 5 embedded fonts (~4.0 MB) | **0 missing `<w:name>`**, font binaries purged, lean **~2.5 MB** (`ZIP_DEFLATED`) |

---

## 2. Five Pillars of Execution

### Pillar 1: Eliminate Useless Blank Spaces & Excess Padding
1. **Body Cleanup**: Iterate all `<w:p>` nodes in the document body (from Section `A. FREE TEST` to the end).
   - Delete any `<w:p>` that contains no text, no drawings (`<w:drawing>`), and no pictures (`<w:pict>`).
   - Delete all whitespace-only paragraphs (`full_text.strip() == ""`).
   - Strip leading and trailing whitespace within text runs (`<w:t>`).
2. **Cover Page Layout Preservation**: Preserve controlled vertical spacing on the cover page (P0–P36) so the title, logo, and student identity blocks remain properly positioned.
3. **Paragraph Spacing Normalization**: Replace manual empty carriage returns with uniform OpenXML paragraph spacing:
   - Body paragraphs: `after="120"` (6 pt), line spacing 1.5 (`line="360"`, `lineRule="auto"`).
   - Paragraph alignment: Justified (`<w:jc w:val="both"/>`).

### Pillar 2: Purge Destructive Page Breaks
1. **Remove Mid-Sentence Page Breaks**:
   - The analysis found page breaks embedded directly inside answers:
     - `P478`: `Jawaban:  Adanya link redundancy dan top...`
     - `P504`: `Jawaban:  Redundancy sudah melekat secar...`
     - `P611`: `Analisis: pemutusan satu kabel PC hanya ...`
   - Strip all `<w:br w:type="page"/>` elements embedded inside paragraph runs across the entire document body.
2. **Remove Isolated Empty Page Break Paragraphs**:
   - Delete empty paragraphs whose only child is a page break (except the designated cover break).
3. **Establish Legitimate Page Boundary**:
   - Retain exactly **one intentional page break** immediately following the cover page footer (`SEPTEMBER 2026`), cleanly separating the cover from Section A.
   - Let Word paginate body content naturally, governed by `keepWithNext` rules.

### Pillar 3: Uniformize All 22 Tables into One Clean Format
1. **Normalize Table Style**:
   - Update every table's `<w:tblStyle>` to `TableGrid`.
   - Ensure `TableGrid` in `word/styles.xml` has valid `<w:name w:val="Table Grid"/>`.
2. **Borders Specification**:
   - Apply a clean, universal border definition to all 22 tables:
     ```xml
     <w:tblBorders>
       <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
       <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
       <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
       <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
       <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
       <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
     </w:tblBorders>
     ```
3. **Cell Margins & Padding**:
   - Enforce comfortable cell margins across all tables: Top: 100 twips, Bottom: 100 twips, Left: 150 twips, Right: 150 twips.
4. **Header Row Formatting & Shading Removal**:
   - Remove dark blue shading (`0a3d66`) and colored fills from header cells. Replace with either a subtle light neutral grey (`F2F2F2`) or clean white background.
   - Ensure header row text is bold and centered/left-aligned.
   - Add `<w:tblHeader/>` to first row so headers repeat properly across multi-page tables.
   - Add `<w:cantSplit/>` to all table rows (`<w:trPr>`) to prevent rows from breaking awkwardly across page splits.
5. **Table Text Typography**:
   - Ensure all text inside table cells uses Times New Roman 11pt, single line spacing, and pure black ink.

### Pillar 4: Total Black Ink Normalization
1. **Remove Colored Run Tags**:
   - Strip all `<w:color>` tags with non-black hex values (`1a1a1a`, `2b2b2b`, `13599c`, `0a3d66`, `555555`, `1c3f1c`, `2d4a22`).
   - Remove `<w:color w:val="ffffff"/>` (white text) originally placed on dark blue table headers, converting them to pure black `#000000`.
2. **Explicit Standard Color**:
   - Ensure runs either inherit automatic black from the document defaults or have `<w:color w:val="000000"/>`.
3. **No Colored Underlines or Accents**:
   - Remove any colored underline attributes (`<w:u w:color="...">`).

### Pillar 5: Automate Heading Detection & Hierarchy
Scan paragraphs using regex patterns and apply formal Word styles and page integrity flags:

1. **Heading 1 (`^[A-Z]\.\s+.*`)**:
   - Matches: `A. FREE TEST - 10 SOAL`, `B. PERTANYAAN URAIAN`, ... `W. POST-TEST` (23 major sections).
   - Style: `Heading 1`
   - Properties: Times New Roman 14pt, Bold, Pure Black, `w:spacing w:before="240" w:after="120"`.
   - Structural Integrity: Inject `<w:keepNext/>` to guarantee headings are never orphaned at the bottom of a page.
2. **Heading 2 (`^\d+\.\s+[A-Z].*`)**:
   - Matches: `1. Pertanyaan Kognitif`, `2. Pertanyaan Afektif`, `1. Gambarkan...` (subsections).
   - Style: `Heading 2`
   - Properties: Times New Roman 12pt, Bold, Pure Black, `w:spacing w:before="160" w:after="80"`.
   - Structural Integrity: Inject `<w:keepNext/>`.
3. **Heading 3 (`^[a-z]\.\s+.*` or sub-numbered items)**:
   - Style: `Heading 3`
   - Properties: Times New Roman 12pt, Bold/Italic, Pure Black.
4. **Body Text**:
   - Normal paragraphs: Times New Roman 12pt, 1.5 line spacing, Justified (`jc="both"`).
   - List items: `List Paragraph` with consistent bullet / number indentation.

### Pillar 6: Layout, Margins & Anti-Bloat Sanitization
1. **Margins**: Set `<w:pgMar>` to standard Indonesian academic 4-4-3-3 cm:
   - Top: `2268` twips (4.00 cm)
   - Left: `2268` twips (4.00 cm)
   - Bottom: `1701` twips (3.00 cm)
   - Right: `1701` twips (3.00 cm)
2. **Paper Size**: Set `<w:pgSz>` to Letter (`w="12240"`, `h="15840"`).
3. **ECMA-376 Compliance**: Ensure all `<w:style>` entries in `word/styles.xml` have valid `<w:name>` tags (0 missing `<w:name>`).
4. **Bloat Prevention**:
   - Strip `<w:embedTrueTypeFonts/>`, `<w:saveSubsetFonts/>`, `<w:doNotEmbedSystemFonts/>` from `word/settings.xml`.
   - Purge `word/fonts/` and `word/_rels/fontTable.xml.rels`.
   - Pack with `zipfile.ZIP_DEFLATED` to ensure clean ~2.5 MB archive.

---

## 3. Step-by-Step Implementation Workflow

```
[Input: documents/Tugas 4 (new).docx]
        │
        ├── Phase 1: Parse OpenXML DOM via lxml
        │       ├── Read document.xml, styles.xml, settings.xml, fontTable.xml
        │
        ├── Phase 2: Structural Sanitation
        │       ├── Purge 229 empty & whitespace paragraphs in body
        │       ├── Eliminate 12 bad/mid-sentence page breaks (keep 1 cover break)
        │       └── Enforce 100% black ink (convert white/blue/grey text to #000000)
        │
        ├── Phase 3: Table Normalization
        │       ├── Apply TableGrid style with 0.5pt single black border
        │       ├── Strip dark blue fills; apply subtle neutral header or white
        │       └── Inject <w:cantSplit/> and <w:tblHeader/>
        │
        ├── Phase 4: Heading Automation
        │       ├── Identify H1 (A-W), H2 (1-18), H3 patterns via regex
        │       └── Apply styles + inject <w:keepNext/>
        │
        ├── Phase 5: Margins & Anti-Bloat
        │       ├── Set 4-4-3-3 cm margins and Letter size across all <w:sectPr>
        │       ├── Strip font embed flags and purge word/fonts/
        │       └── Repack via ZIP_DEFLATED
        │
        └── Phase 6: Automated Verification & User Validation
```

---

## 4. Opinion & Recommendations for a Truly Uniform, Clean Document

Here is my expert perspective on the key elements required to achieve a pristine, publication-grade academic document:

### 1. Ban Manual Carriage Returns for Spacing
The root cause of "cluttered" documents is pressing Enter multiple times to create vertical space. This creates brittle documents where adding one sentence pushes headings onto empty pages.
* **Best Practice**: Set `w:spacing w:after="120"` (6pt) on paragraphs and `w:spacing w:before="240"` (12pt) on headings. The spacing is mathematically consistent, and the document stays compact and clean.

### 2. Mandatory `<w:keepNext/>` on Every Heading
Nothing looks more unprofessional than a section title (e.g., `B. PERTANYAAN URAIAN`) sitting abandoned as the very last line at the bottom of a page, with its question appearing on the next page.
* **Best Practice**: Every automated heading (H1, H2, H3) **must** have `<w:keepNext/>` in its paragraph properties. Word will automatically move the heading to the next page if the first paragraph following it doesn't fit.

### 3. Mandatory `<w:cantSplit/>` on Table Rows
In technical reports with large question/answer tables, Word frequently cuts a single table row in half across a page boundary (showing top half of a text block on page 5 and bottom half on page 6).
* **Best Practice**: Inject `<w:cantSplit/>` into every `<w:trPr>`. This forces entire rows to move together, dramatically enhancing readability.

### 4. Minimalist Neutral Tables vs. Heavy Shading
The dark navy blue shading (`0a3d66`) with white text found in some tables of `Tugas 4 (new).docx` creates jarring visual discord against regular text and prints poorly in black-and-white.
* **Best Practice**: Use standard black gridlines (`0.5pt`) with either a clean white cell background or an ultra-light neutral grey header (`#F2F2F2`) with bold black text. This matches standard academic IEEE/ACM and Indonesian thesis guidelines.

### 5. Clear Heading Hierarchy (H1 vs H2 vs H3)
Currently, `Tugas 4 (new).docx` relies on ad-hoc bold text. Promoting major sections (`A.` through `W.`) to true Word `Heading 1` and numbered subsections to `Heading 2` allows Word to generate an instant, automated Table of Contents (*Daftar Isi*) and enables the Navigation Pane for immediate section jumping.
