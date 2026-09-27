### 1. Root Cause Breakdown: What Actually Happened

#### A. The File Bloat (3.4 MB $\to$ 23.3 MB)

The 20 MB size explosion was driven by two independent factors in Word's OpenXML packaging:

* **Uncompressed Font Binary Dumps (+11.3 MB compressed / +22.7 MB uncompressed):**
* `word/settings.xml` contained `<w:embedTrueTypeFonts/>`, `<w:doNotEmbedSystemFonts w:val="false"/>`, and `<w:saveSubsetFonts w:val="false"/>`.
* Because subsetting was disabled (`saveSubsetFonts=false`), Word did not extract only the characters used in your document. It dumped full binaries of every referenced system font directly into `word/fonts/`.
* The single largest asset was `font6.odttf` (Gungsuh, a full CJK Korean font binary), totaling **16.1 MB** uncompressed. Standard fonts (Calibri, Cambria, Times) added another ~6 MB.


* **Storage Compression Failure (+9.3 MB):**
* In the original backup, all 39 embedded images in `word/media/` were compressed using standard `ZIP_DEFLATED` (2.55 MB total).
* Word re-saved these 39 images with `compress_type: 0` (`ZIP_STORED`), writing them to disk as uncompressed raw bitmaps totaling **11.87 MB**.



#### B. The "Styles 1" Corruption Prompt

* Under OpenXML standard ECMA-376 (ISO/IEC 29500), every `<w:style>` node must contain a `<w:name>` child element as its primary metadata descriptor.


* The document originated from a non-standard generator (Google Docs or WPS export) that left **24 malformed table style entries** (`a`, `a0` through `af6`) lacking `<w:name>` entirely.


* Microsoft Word's parser executes strict schema validation on open. When encountering a `<w:style>` lacking `<w:name>`, it flags the style catalog as invalid, terminates parsing, triggers the **"Show Repairs: Styles 1"** recovery prompt, and discards those styles.



---

### 2. Anatomy of the Overcomplication

The execution derailed into iterative confusion due to three fundamental engineering mistakes:

```
[Problem: 24 Styles Missing <w:name>]
        │
        ├── Wrong Path (The Agent): "Transplant entire styles.xml & rewrite document.xml"
        │       └── Result: High blast radius, lost styles, broken table borders.
        │
        ├── Contradictory Path: "Hybrid fix without touching document.xml"
        │       └── Result: Impossible. Margins (<w:pgMar>) only exist in document.xml.
        │
        └── Correct Path (Surgical): "Inject <w:name> directly into the 24 nodes in-place"
                └── Result: Zero blast radius, 100% schema compliant, zero doc body edits.

```

#### Mistake 1: The "Replace Everything" Fallacy

Rather than isolating the 24 missing tags, the initial strategy attempted a wholesale transplant of `word/styles.xml` from `EVALUASI MATERI 3 (1).docx` into `Tugas 4.docx`.

* **The Weak Point:** Style catalogs in OpenXML are tightly coupled to the document body. Swapping style sheets assumes identical GUIDs and style names across both documents.


* **The Collateral Damage:** It required executing regex replacements across `word/document.xml` to force all table references to `TableGrid`. This would strip custom cell margins, borderless formatting, and custom padding across all 24 tables.



#### Mistake 2: The Architectural Impossibility (Margins Claim)

The agent claimed it could implement the "Hybrid Strategy" (applying 4-4-3-3 cm margins) *"without touching `document.xml`"*.

* **The Reality:** OpenXML specifies page margins (`<w:pgMar>`) exclusively within section properties (`<w:sectPr>`). Section properties reside inside `<w:body>` at the bottom of `word/document.xml`. They **cannot** be defined or overridden in `word/styles.xml`. Claiming margins can be set without modifying `document.xml` demonstrated a misunderstanding of OpenXML DOM hierarchy.

#### Mistake 3: PowerShell Scripting Fragility

Running inline multi-line Python scripts via PowerShell using double quotes (`python -c "..."`) caused parser collisions on nested quotes. Instead of recognizing the CLI escaping failure, execution stalled on debugging command syntax rather than resolving document logic.

---

### 3. Comparison of Repair Strategies

| Evaluation Dimension | Full Transplant

 | Proposed Hybrid (Flawed Claim)

 | Corrected Surgical Architecture |
| --- | --- | --- | --- |
| **"Styles 1" Error Fix** | Yes

 | Yes

 | **Yes (Directly satisfies ECMA-376)**<br> |
| **`document.xml` Modification** | Heavy (Rewrites all table styles)

 | Claims None (Factually impossible)

 | **Minimal Surgical** (Modifies only `<w:pgMar>`, leaves paragraphs/tables untouched) |
| **Table Formatting Risk** | High (Forces generic `TableGrid` borders on all 24 tables)

 | Zero

 | **Zero (Existing table properties remain 100% preserved)**<br> |
| **Page Margin Application** | Applied

 | Failed / Unachievable

 | **Applied (2268 / 1701 twips explicitly set)** |
| **Font Bloat Elimination** | Incomplete without settings purge | Incomplete without settings purge | **Complete (Deletes `/fonts/` and unsets embed flags)** |
| **Final File Size** | ~2.7 MB

 | ~2.7 MB

 | **~2.7 MB (Recompressed with `ZIP_DEFLATED`)** |

---

### 4. Hidden Technical Traps to Guard Against

* **The Word Re-Save Trap:** If you merely delete the `word/fonts/` folder from the ZIP archive without stripping `<w:embedTrueTypeFonts/>` from `word/settings.xml`, the very next time you open the document in Word and press `Ctrl+S`, Word will scan your operating system, pull the 16 MB Korean font binary again, and bloat the file back to 23 MB.
* **Inline Run Formatting Overrides:** Setting Times New Roman 12pt inside `<w:style w:styleId="Normal">` only updates paragraphs that strictly inherit from `Normal`. Any paragraph or run where the original author manually applied inline font tags (`<w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr>`) will override the `Normal` style.
* **Twip Precision:** OpenXML margins use twips (twentieth of an imperial point; $1\text{ cm} \approx 567\text{ twips}$). Setting 4-4-3-3 cm requires exact values:
* Top: $4\text{ cm} \times 567 = 2268\text{ twips}$
* Left: $4\text{ cm} \times 567 = 2268\text{ twips}$
* Bottom: $3\text{ cm} \times 567 = 1701\text{ twips}$
* Right: $3\text{ cm} \times 567 = 1701\text{ twips}$



---

### 5. Definitive Verdict & Architectural Specification

**Abandon all full-transplant or file-swapping ideas.** The only clean, risk-free path is a multi-part surgical XML update executed within the ZIP container:

1. **`word/styles.xml`:**
* Iterate all `<w:style>` nodes. Where `<w:name>` is absent, inject `<w:name w:val="Table {styleId}"/>` at index 0.


* Mutate the existing `Normal` style node to point to Times New Roman at 12pt (`sz=24`).


2. **`word/document.xml`:**
* Traverse only to the root `<w:body>/<w:sectPr>/<w:pgMar>` element.
* Overwrite only the four margin attributes (`top="2268"`, `left="2268"`, `bottom="1701"`, `right="1701"`).
* **Do not touch** any `<w:p>` (paragraph) or `<w:tbl>` (table) nodes.


3. **`word/settings.xml` & Package Structure:**
* Strip `<w:embedTrueTypeFonts/>`, `<w:saveSubsetFonts/>`, and `<w:doNotEmbedSystemFonts/>` completely.
* Delete `word/fonts/` entirely from the ZIP manifest.
* Pack all surviving XML and media files using standard `ZIP_DEFLATED` compression.



This specification guarantees 100% ECMA-376 schema compliance, zero table layout corruption, precise 4-4-3-3 cm margins, and permanent debloating to ~2.7 MB.