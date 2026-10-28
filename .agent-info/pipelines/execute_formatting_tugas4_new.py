import zipfile
import shutil
import os
import re
from lxml import etree
import docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def ensure_child(parent, tag_name):
    child = parent.find(f'w:{tag_name}', namespaces=NS)
    if child is None:
        child = etree.SubElement(parent, f'{{{W_NS}}}{tag_name}')
    return child

def format_document(src_path, dst_path):
    print(f"=== PROCESSING DOCUMENT ===")
    print(f"Source: {src_path}")
    print(f"Destination: {dst_path}")
    
    with zipfile.ZipFile(src_path, 'r') as zin:
        namelist = zin.namelist()
        
        # -------------------------------------------------------------
        # 1. PROCESS word/styles.xml
        # -------------------------------------------------------------
        styles_xml = etree.fromstring(zin.read('word/styles.xml'))
        
        # Ensure all existing styles have valid <w:name>
        for s in styles_xml.findall('w:style', namespaces=NS):
            if s.find('w:name', namespaces=NS) is None:
                sid = s.attrib.get(f'{{{W_NS}}}styleId', 'Unknown')
                new_name = etree.Element(f'{{{W_NS}}}name')
                new_name.attrib[f'{{{W_NS}}}val'] = f'Table {sid}'
                s.insert(0, new_name)
                
        def get_or_create_style(style_id, style_type, name_val):
            for s in styles_xml.findall('w:style', namespaces=NS):
                if s.attrib.get(f'{{{W_NS}}}styleId') == style_id:
                    return s
            s = etree.SubElement(styles_xml, f'{{{W_NS}}}style')
            s.attrib[f'{{{W_NS}}}type'] = style_type
            s.attrib[f'{{{W_NS}}}styleId'] = style_id
            name_el = etree.SubElement(s, f'{{{W_NS}}}name')
            name_el.attrib[f'{{{W_NS}}}val'] = name_val
            return s
            
        # Normal style (Times New Roman 12pt, 1.5 line, justified)
        normal_s = get_or_create_style('Normal', 'paragraph', 'Normal')
        pPr = ensure_child(normal_s, 'pPr')
        spacing = ensure_child(pPr, 'spacing')
        spacing.attrib[f'{{{W_NS}}}after'] = '120'
        spacing.attrib[f'{{{W_NS}}}line'] = '360'
        spacing.attrib[f'{{{W_NS}}}lineRule'] = 'auto'
        jc = ensure_child(pPr, 'jc')
        jc.attrib[f'{{{W_NS}}}val'] = 'both'
        
        rPr = ensure_child(normal_s, 'rPr')
        rFonts = ensure_child(rPr, 'rFonts')
        rFonts.attrib[f'{{{W_NS}}}ascii'] = 'Times New Roman'
        rFonts.attrib[f'{{{W_NS}}}hAnsi'] = 'Times New Roman'
        rFonts.attrib[f'{{{W_NS}}}cs'] = 'Times New Roman'
        ensure_child(rPr, 'sz').attrib[f'{{{W_NS}}}val'] = '24'
        ensure_child(rPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '24'
        ensure_child(rPr, 'color').attrib[f'{{{W_NS}}}val'] = '000000'
        
        # Heading 1 (Times New Roman 14pt, Bold, black)
        h1_s = get_or_create_style('Heading1', 'paragraph', 'heading 1')
        pPr1 = ensure_child(h1_s, 'pPr')
        if pPr1.find('w:keepNext', namespaces=NS) is None:
            etree.SubElement(pPr1, f'{{{W_NS}}}keepNext')
        sp1 = ensure_child(pPr1, 'spacing')
        sp1.attrib[f'{{{W_NS}}}before'] = '240'
        sp1.attrib[f'{{{W_NS}}}after'] = '120'
        rPr1 = ensure_child(h1_s, 'rPr')
        rf1 = ensure_child(rPr1, 'rFonts')
        rf1.attrib[f'{{{W_NS}}}ascii'] = 'Times New Roman'
        rf1.attrib[f'{{{W_NS}}}hAnsi'] = 'Times New Roman'
        if rPr1.find('w:b', namespaces=NS) is None:
            etree.SubElement(rPr1, f'{{{W_NS}}}b')
        if rPr1.find('w:bCs', namespaces=NS) is None:
            etree.SubElement(rPr1, f'{{{W_NS}}}bCs')
        ensure_child(rPr1, 'sz').attrib[f'{{{W_NS}}}val'] = '28'
        ensure_child(rPr1, 'szCs').attrib[f'{{{W_NS}}}val'] = '28'
        ensure_child(rPr1, 'color').attrib[f'{{{W_NS}}}val'] = '000000'

        # Heading 2 (Times New Roman 12pt, Bold, black)
        h2_s = get_or_create_style('Heading2', 'paragraph', 'heading 2')
        pPr2 = ensure_child(h2_s, 'pPr')
        if pPr2.find('w:keepNext', namespaces=NS) is None:
            etree.SubElement(pPr2, f'{{{W_NS}}}keepNext')
        sp2 = ensure_child(pPr2, 'spacing')
        sp2.attrib[f'{{{W_NS}}}before'] = '160'
        sp2.attrib[f'{{{W_NS}}}after'] = '80'
        rPr2 = ensure_child(h2_s, 'rPr')
        rf2 = ensure_child(rPr2, 'rFonts')
        rf2.attrib[f'{{{W_NS}}}ascii'] = 'Times New Roman'
        rf2.attrib[f'{{{W_NS}}}hAnsi'] = 'Times New Roman'
        if rPr2.find('w:b', namespaces=NS) is None:
            etree.SubElement(rPr2, f'{{{W_NS}}}b')
        if rPr2.find('w:bCs', namespaces=NS) is None:
            etree.SubElement(rPr2, f'{{{W_NS}}}bCs')
        ensure_child(rPr2, 'sz').attrib[f'{{{W_NS}}}val'] = '24'
        ensure_child(rPr2, 'szCs').attrib[f'{{{W_NS}}}val'] = '24'
        ensure_child(rPr2, 'color').attrib[f'{{{W_NS}}}val'] = '000000'

        # Ensure TableGrid exists
        get_or_create_style('TableGrid', 'table', 'Table Grid')
        
        # -------------------------------------------------------------
        # 2. PROCESS word/document.xml
        # -------------------------------------------------------------
        doc_xml = etree.fromstring(zin.read('word/document.xml'))
        body = doc_xml.find('w:body', namespaces=NS)
        
        # Find cover page boundary (paragraph containing 'SEPTEMBER 2026')
        all_children = list(body)
        cover_end_idx = 36
        for i, child in enumerate(all_children):
            if child.tag.endswith('p'):
                text = "".join(t.text for t in child.findall('.//w:t', namespaces=NS) if t.text).strip()
                if 'SEPTEMBER 2026' in text.upper():
                    cover_end_idx = i
                    break
        print(f"Cover page boundary detected at index: {cover_end_idx}")

        # Strip all existing page breaks across the whole document body
        deleted_bad_br = 0
        for br in body.findall('.//w:br[@w:type="page"]', namespaces=NS):
            p_parent = br.getparent()
            p_parent.remove(br)
            deleted_bad_br += 1

        # Delete useless empty paragraphs in the body (after cover page)
        new_children = []
        deleted_empty_p = 0
        for i, child in enumerate(all_children):
            if child.tag.endswith('sectPr'):
                continue
                
            if i > cover_end_idx and child.tag.endswith('p'):
                text = "".join(t.text for t in child.findall('.//w:t', namespaces=NS) if t.text is not None).strip()
                drawings = child.findall('.//w:drawing', namespaces=NS)
                picts = child.findall('.//w:pict', namespaces=NS)
                if not text and len(drawings) == 0 and len(picts) == 0:
                    deleted_empty_p += 1
                    continue # omit empty paragraph
            
            # Place 1 intentional clean page break right after the cover page
            if i == cover_end_idx:
                new_children.append(child)
                pb_p = etree.Element(f'{{{W_NS}}}p')
                pb_r = etree.SubElement(pb_p, f'{{{W_NS}}}r')
                pb_br = etree.SubElement(pb_r, f'{{{W_NS}}}br')
                pb_br.attrib[f'{{{W_NS}}}type'] = 'page'
                new_children.append(pb_p)
                continue
                
            new_children.append(child)

        root_sectPr = body.find('w:sectPr', namespaces=NS)
        body.clear()
        for child in new_children:
            body.append(child)
        if root_sectPr is not None:
            body.append(root_sectPr)

        print(f"Deleted useless empty/whitespace paragraphs in body: {deleted_empty_p}")
        print(f"Deleted bad page breaks: {deleted_bad_br} (1 clean page break retained after cover)")

        # -------------------------------------------------------------
        # 3. 100% BLACK INK NORMALIZATION ACROSS ENTIRE DOCUMENT
        # -------------------------------------------------------------
        color_fixed = 0
        for c in doc_xml.findall('.//w:color', namespaces=NS):
            val = c.attrib.get(f'{{{W_NS}}}val', '')
            if val.lower() != '000000':
                c.attrib[f'{{{W_NS}}}val'] = '000000'
                color_fixed += 1
        print(f"Normalized {color_fixed} non-black color tags to 100% black ink (000000)")

        # -------------------------------------------------------------
        # 4. AUTOMATED HEADING DETECTION & STYLING
        # -------------------------------------------------------------
        h1_regex = re.compile(r'^[A-Z]\.\s+[A-Z0-9].*')
        h1_blacklist_terms = ["A. PROTOKOL", "A. BUS", "A. STAR", "A. 2B.", "A. SEMUA NODE", "A. TIDAK MEMILIKI", "A. BACKBONE"]
        h2_regex = re.compile(r'^\d+\.\s+[A-Z].*')

        automated_h1 = 0
        automated_h2 = 0
        for p in body.findall('.//w:p', namespaces=NS):
            text = "".join(t.text for t in p.findall('.//w:t', namespaces=NS) if t.text is not None).strip()
            if not text:
                continue
                
            pPr = ensure_child(p, 'pPr')
            is_h1 = False
            if h1_regex.match(text) and len(text) < 120:
                if not any(term in text.upper() for term in h1_blacklist_terms):
                    is_h1 = True
                    
            is_h2 = False
            if not is_h1 and h2_regex.match(text) and len(text) < 90:
                if not re.match(r'^\d+\.\s+[A-E]\b', text):
                    is_h2 = True

            if is_h1:
                pStyle = ensure_child(pPr, 'pStyle')
                pStyle.attrib[f'{{{W_NS}}}val'] = 'Heading1'
                if pPr.find('w:keepNext', namespaces=NS) is None:
                    pPr.insert(0, etree.Element(f'{{{W_NS}}}keepNext'))
                automated_h1 += 1
            elif is_h2:
                pStyle = ensure_child(pPr, 'pStyle')
                pStyle.attrib[f'{{{W_NS}}}val'] = 'Heading2'
                if pPr.find('w:keepNext', namespaces=NS) is None:
                    pPr.insert(0, etree.Element(f'{{{W_NS}}}keepNext'))
                automated_h2 += 1
            else:
                # Body paragraph properties
                spacing = ensure_child(pPr, 'spacing')
                if f'{{{W_NS}}}after' not in spacing.attrib:
                    spacing.attrib[f'{{{W_NS}}}after'] = '120'
                if f'{{{W_NS}}}line' not in spacing.attrib:
                    spacing.attrib[f'{{{W_NS}}}line'] = '360'
                    spacing.attrib[f'{{{W_NS}}}lineRule'] = 'auto'
                jc = ensure_child(pPr, 'jc')
                if f'{{{W_NS}}}val' not in jc.attrib:
                    jc.attrib[f'{{{W_NS}}}val'] = 'both'

        print(f"Automated Heading 1: {automated_h1} (with keepNext)")
        print(f"Automated Heading 2: {automated_h2} (with keepNext)")

        # -------------------------------------------------------------
        # 5. UNIFORMIZE ALL 22 TABLES
        # -------------------------------------------------------------
        tables = doc_xml.findall('.//w:tbl', namespaces=NS)
        print(f"Uniformizing {len(tables)} tables...")
        for tbl in tables:
            tblPr = ensure_child(tbl, 'tblPr')
            
            # Remove existing tblStyle elements and add single TableGrid
            for old_style in tblPr.findall('w:tblStyle', namespaces=NS):
                tblPr.remove(old_style)
            tblStyle = etree.Element(f'{{{W_NS}}}tblStyle')
            tblStyle.attrib[f'{{{W_NS}}}val'] = 'TableGrid'
            tblPr.insert(0, tblStyle)
            
            # Remove existing borders and add clean 0.5pt black grid
            for old_b in tblPr.findall('w:tblBorders', namespaces=NS):
                tblPr.remove(old_b)
            borders = etree.SubElement(tblPr, f'{{{W_NS}}}tblBorders')
            for b_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                b_el = etree.SubElement(borders, f'{{{W_NS}}}{b_name}')
                b_el.attrib[f'{{{W_NS}}}val'] = 'single'
                b_el.attrib[f'{{{W_NS}}}sz'] = '4'
                b_el.attrib[f'{{{W_NS}}}space'] = '0'
                b_el.attrib[f'{{{W_NS}}}color'] = '000000'
                
            # Cell Margins / Padding
            for old_mar in tblPr.findall('w:tblCellMar', namespaces=NS):
                tblPr.remove(old_mar)
            cellMar = etree.SubElement(tblPr, f'{{{W_NS}}}tblCellMar')
            for m_name, m_val in [('top', '100'), ('bottom', '100'), ('left', '150'), ('right', '150')]:
                m_el = etree.SubElement(cellMar, f'{{{W_NS}}}{m_name}')
                m_el.attrib[f'{{{W_NS}}}w'] = m_val
                m_el.attrib[f'{{{W_NS}}}type'] = 'dxa'
                
            # Rows processing
            rows = tbl.findall('w:tr', namespaces=NS)
            for r_idx, row in enumerate(rows):
                trPr = ensure_child(row, 'trPr')
                # cantSplit on every row
                if trPr.find('w:cantSplit', namespaces=NS) is None:
                    etree.SubElement(trPr, f'{{{W_NS}}}cantSplit')
                # tblHeader ONLY on row 0
                if r_idx == 0:
                    if trPr.find('w:tblHeader', namespaces=NS) is None:
                        etree.SubElement(trPr, f'{{{W_NS}}}tblHeader')
                else:
                    for old_th in trPr.findall('w:tblHeader', namespaces=NS):
                        trPr.remove(old_th)
                        
                # Cells processing
                for cell in row.findall('w:tc', namespaces=NS):
                    tcPr = ensure_child(cell, 'tcPr')
                    # Shading: F2F2F2 for header, clear for data
                    shd = ensure_child(tcPr, 'shd')
                    shd.attrib[f'{{{W_NS}}}val'] = 'clear'
                    shd.attrib[f'{{{W_NS}}}color'] = 'auto'
                    if r_idx == 0:
                        shd.attrib[f'{{{W_NS}}}fill'] = 'F2F2F2'
                    else:
                        shd.attrib[f'{{{W_NS}}}fill'] = 'FFFFFF'
                        
                    for cp in cell.findall('.//w:p', namespaces=NS):
                        for cr in cp.findall('.//w:r', namespaces=NS):
                            crPr = ensure_child(cr, 'rPr')
                            if r_idx == 0:
                                if crPr.find('w:b', namespaces=NS) is None:
                                    etree.SubElement(crPr, f'{{{W_NS}}}b')
                            ensure_child(crPr, 'color').attrib[f'{{{W_NS}}}val'] = '000000'
                            crf = ensure_child(crPr, 'rFonts')
                            crf.attrib[f'{{{W_NS}}}ascii'] = 'Times New Roman'
                            crf.attrib[f'{{{W_NS}}}hAnsi'] = 'Times New Roman'
                            ensure_child(crPr, 'sz').attrib[f'{{{W_NS}}}val'] = '22'
                            ensure_child(crPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '22'

        # -------------------------------------------------------------
        # 6. SECTION MARGINS (4-4-3-3 cm) & LETTER SIZE
        # -------------------------------------------------------------
        sectPr_count = 0
        for sectPr in doc_xml.findall('.//w:sectPr', namespaces=NS):
            sectPr_count += 1
            pgMar = ensure_child(sectPr, 'pgMar')
            pgMar.attrib[f'{{{W_NS}}}top'] = '2268'     # 4.00 cm
            pgMar.attrib[f'{{{W_NS}}}left'] = '2268'    # 4.00 cm
            pgMar.attrib[f'{{{W_NS}}}bottom'] = '1701'  # 3.00 cm
            pgMar.attrib[f'{{{W_NS}}}right'] = '1701'   # 3.00 cm
            pgMar.attrib[f'{{{W_NS}}}header'] = '720'
            pgMar.attrib[f'{{{W_NS}}}footer'] = '720'
            
            pgSz = ensure_child(sectPr, 'pgSz')
            pgSz.attrib[f'{{{W_NS}}}w'] = '12240'       # Letter width
            pgSz.attrib[f'{{{W_NS}}}h'] = '15840'       # Letter height
            pgSz.attrib[f'{{{W_NS}}}orient'] = 'portrait'
        print(f"Updated {sectPr_count} section(s) to 4-4-3-3 cm margins and Letter size")

        # -------------------------------------------------------------
        # 7. STRIP FONT EMBED SETTINGS & REFERENCES
        # -------------------------------------------------------------
        settings_xml = etree.fromstring(zin.read('word/settings.xml'))
        for el in list(settings_xml):
            tag = el.tag.split('}')[-1]
            if tag in ['embedTrueTypeFonts', 'saveSubsetFonts', 'doNotEmbedSystemFonts']:
                settings_xml.remove(el)

        fonttable_xml = etree.fromstring(zin.read('word/fontTable.xml'))
        for font in fonttable_xml.findall('w:font', namespaces=NS):
            for child in list(font):
                tag = child.tag.split('}')[-1]
                if tag.startswith('embed'):
                    font.remove(child)

        # -------------------------------------------------------------
        # 8. REPACK WITH ZIP_DEFLATED
        # -------------------------------------------------------------
        tmp_dst = dst_path + ".tmp"
        with zipfile.ZipFile(tmp_dst, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in namelist:
                if item.startswith('word/fonts/') or item == 'word/_rels/fontTable.xml.rels':
                    continue
                if item == 'word/styles.xml':
                    zout.writestr(item, etree.tostring(styles_xml, xml_declaration=True, encoding='utf-8', standalone=True))
                elif item == 'word/document.xml':
                    zout.writestr(item, etree.tostring(doc_xml, xml_declaration=True, encoding='utf-8', standalone=True))
                elif item == 'word/settings.xml':
                    zout.writestr(item, etree.tostring(settings_xml, xml_declaration=True, encoding='utf-8', standalone=True))
                elif item == 'word/fontTable.xml':
                    zout.writestr(item, etree.tostring(fonttable_xml, xml_declaration=True, encoding='utf-8', standalone=True))
                else:
                    data = zin.read(item)
                    zout.writestr(item, data)

        if os.path.exists(dst_path):
            os.remove(dst_path)
        os.rename(tmp_dst, dst_path)

    print(f"\n==================================================")
    print(f"SUCCESS: Formatted document written to {dst_path}")
    print(f"Original size: {os.path.getsize(src_path) / (1024*1024):.2f} MB")
    print(f"Final size:    {os.path.getsize(dst_path) / (1024*1024):.2f} MB")
    print(f"==================================================")

if __name__ == '__main__':
    src = 'documents/Tugas 4 (new)_backup.docx'
    dst = 'documents/Tugas 4 (new).docx'
    format_document(src, dst)
