import zipfile
import shutil
import os
import re
from lxml import etree
import docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

import io

def ensure_child(parent, tag_name):
    child = parent.find(f'w:{tag_name}', namespaces=NS)
    if child is None:
        child = etree.SubElement(parent, f'{{{W_NS}}}{tag_name}')
    return child

def apply_refinements(src_path, dst_path):
    print(f"=== APPLYING REFINEMENTS ===")
    print(f"Source: {src_path}")
    print(f"Target: {dst_path}")
    
    with open(src_path, 'rb') as f:
        src_bytes = io.BytesIO(f.read())
        
    with zipfile.ZipFile(src_bytes, 'r') as zin:
        namelist = zin.namelist()
        
        # -------------------------------------------------------------
        # 1. PROCESS word/styles.xml
        # -------------------------------------------------------------
        styles_xml = etree.fromstring(zin.read('word/styles.xml'))
        
        # Update Heading 1 to 12pt Bold (per user request!)
        for s in styles_xml.findall('w:style', namespaces=NS):
            sid = s.attrib.get(f'{{{W_NS}}}styleId')
            if sid == 'Heading1':
                rPr = ensure_child(s, 'rPr')
                ensure_child(rPr, 'sz').attrib[f'{{{W_NS}}}val'] = '24'     # 12 pt
                ensure_child(rPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '24'   # 12 pt
                ensure_child(rPr, 'color').attrib[f'{{{W_NS}}}val'] = '000000'
                if rPr.find('w:b', namespaces=NS) is None:
                    etree.SubElement(rPr, f'{{{W_NS}}}b')
                if rPr.find('w:bCs', namespaces=NS) is None:
                    etree.SubElement(rPr, f'{{{W_NS}}}bCs')
            elif sid == 'Heading2':
                rPr = ensure_child(s, 'rPr')
                ensure_child(rPr, 'sz').attrib[f'{{{W_NS}}}val'] = '24'     # 12 pt
                ensure_child(rPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '24'   # 12 pt
                ensure_child(rPr, 'color').attrib[f'{{{W_NS}}}val'] = '000000'
                if rPr.find('w:b', namespaces=NS) is None:
                    etree.SubElement(rPr, f'{{{W_NS}}}b')
                if rPr.find('w:bCs', namespaces=NS) is None:
                    etree.SubElement(rPr, f'{{{W_NS}}}bCs')
                    
        # -------------------------------------------------------------
        # 2. PROCESS word/document.xml
        # -------------------------------------------------------------
        doc_xml = etree.fromstring(zin.read('word/document.xml'))
        body = doc_xml.find('w:body', namespaces=NS)
        
        # Detect cover boundary
        all_children = list(body)
        cover_end_idx = 36
        for i, child in enumerate(all_children):
            if child.tag.endswith('p'):
                t = "".join(node.text for node in child.findall('.//w:t', namespaces=NS) if node.text).strip()
                if 'SEPTEMBER 2026' in t.upper():
                    cover_end_idx = i
                    break

        # A. TABLE SANITIZATION (Purge empty Table 9, Extract images from Table 1 & Table 3)
        deleted_tables = 0
        extracted_tables = 0
        
        for tbl in list(body.findall('w:tbl', namespaces=NS)):
            text_nodes = tbl.findall('.//w:t', namespaces=NS)
            full_text = "".join(t.text for t in text_nodes if t.text).strip()
            drawings = tbl.findall('.//w:drawing', namespaces=NS)
            picts = tbl.findall('.//w:pict', namespaces=NS)
            total_images = len(drawings) + len(picts)
            
            # Case 1: Empty table holding no value
            if not full_text and total_images == 0:
                body.remove(tbl)
                deleted_tables += 1
                print(f"Deleted empty table holding no value")
                continue
                
            # Case 2: Image container table (images only, text < 15)
            if total_images > 0 and len(full_text) < 15:
                # Create a clean centered paragraph for each drawing
                tbl_parent = tbl.getparent()
                tbl_pos = tbl_parent.index(tbl)
                
                for d in drawings:
                    img_p = etree.Element(f'{{{W_NS}}}p')
                    img_pPr = etree.SubElement(img_p, f'{{{W_NS}}}pPr')
                    jc = etree.SubElement(img_pPr, f'{{{W_NS}}}jc')
                    jc.attrib[f'{{{W_NS}}}val'] = 'center'
                    sp = etree.SubElement(img_pPr, f'{{{W_NS}}}spacing')
                    sp.attrib[f'{{{W_NS}}}before'] = '160'
                    sp.attrib[f'{{{W_NS}}}after'] = '120'
                    etree.SubElement(img_pPr, f'{{{W_NS}}}keepNext')
                    
                    img_r = etree.SubElement(img_p, f'{{{W_NS}}}r')
                    img_r.append(d)
                    
                    tbl_parent.insert(tbl_pos, img_p)
                    tbl_pos += 1
                    
                for pi in picts:
                    img_p = etree.Element(f'{{{W_NS}}}p')
                    img_pPr = etree.SubElement(img_p, f'{{{W_NS}}}pPr')
                    jc = etree.SubElement(img_pPr, f'{{{W_NS}}}jc')
                    jc.attrib[f'{{{W_NS}}}val'] = 'center'
                    etree.SubElement(img_pPr, f'{{{W_NS}}}keepNext')
                    img_r = etree.SubElement(img_p, f'{{{W_NS}}}r')
                    img_r.append(pi)
                    tbl_parent.insert(tbl_pos, img_p)
                    tbl_pos += 1
                    
                tbl_parent.remove(tbl)
                extracted_tables += 1
                print(f"Extracted image(s) from container table and deleted table")
                continue

        print(f"Tables result: Deleted {deleted_tables} empty, Extracted {extracted_tables} image-containers.")
        surviving_tables = doc_xml.findall('.//w:tbl', namespaces=NS)
        print(f"Surviving data tables: {len(surviving_tables)}")

        # B. MULTIPLE CHOICE LEFT-ALIGNMENT & ABNORMAL BOLD CLEANUP & FONT SIZES
        mc_regex = re.compile(r'^[A-E]\.\s+')
        multi_mc_regex = re.compile(r'A\.\s+.*B\.\s+.*C\.\s+')
        
        # Lead-in label patterns (including bulleted items like • Biaya:, • Performa:)
        lead_in_regex = re.compile(r'^([•\-\*]?\s*[A-Za-z0-9\s/–\-]{2,35}:)\s*(.*)', re.IGNORECASE)
        
        unbolded_paragraphs = 0
        mc_aligned = 0
        
        for p_idx, p in enumerate(body.findall('.//w:p', namespaces=NS)):
            pPr = ensure_child(p, 'pPr')
            pStyle_el = pPr.find('w:pStyle', namespaces=NS)
            pStyle = pStyle_el.attrib.get(f'{{{W_NS}}}val') if pStyle_el is not None else None
            
            full_text = "".join(t.text for t in p.findall('.//w:t', namespaces=NS) if t.text is not None).strip()
            if not full_text:
                continue

            # 1. Multiple Choice Alignment Exemption (Left-align)
            is_mc = False
            if mc_regex.match(full_text) or multi_mc_regex.search(full_text) or (full_text.startswith("A.") and "B." in full_text):
                is_mc = True
                jc = ensure_child(pPr, 'jc')
                jc.attrib[f'{{{W_NS}}}val'] = 'left'
                mc_aligned += 1

            # 2. Heading 1 Font Size (12 pt per user request)
            if pStyle == 'Heading1':
                for r in p.findall('.//w:r', namespaces=NS):
                    rPr = ensure_child(r, 'rPr')
                    ensure_child(rPr, 'sz').attrib[f'{{{W_NS}}}val'] = '24'   # 12 pt
                    ensure_child(rPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '24' # 12 pt
                    if rPr.find('w:b', namespaces=NS) is None:
                        etree.SubElement(rPr, f'{{{W_NS}}}b')
                    if rPr.find('w:bCs', namespaces=NS) is None:
                        etree.SubElement(rPr, f'{{{W_NS}}}bCs')
                continue # Don't unbold headings

            # 3. Heading 2 Font Size (12 pt)
            if pStyle == 'Heading2':
                for r in p.findall('.//w:r', namespaces=NS):
                    rPr = ensure_child(r, 'rPr')
                    ensure_child(rPr, 'sz').attrib[f'{{{W_NS}}}val'] = '24'   # 12 pt
                    ensure_child(rPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '24' # 12 pt
                    if rPr.find('w:b', namespaces=NS) is None:
                        etree.SubElement(rPr, f'{{{W_NS}}}b')
                    if rPr.find('w:bCs', namespaces=NS) is None:
                        etree.SubElement(rPr, f'{{{W_NS}}}bCs')
                continue # Don't unbold headings

            # Skip cover page paragraphs
            if p_idx <= cover_end_idx:
                continue

            # 4. Check if paragraph is inside a table
            in_table = bool(p.xpath('ancestor::w:tbl', namespaces=NS))
            if in_table:
                # Inside tables: ensure 11pt
                for r in p.findall('.//w:r', namespaces=NS):
                    rPr = ensure_child(r, 'rPr')
                    ensure_child(rPr, 'sz').attrib[f'{{{W_NS}}}val'] = '22'
                    ensure_child(rPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '22'
                continue

            # 5. Normal Body Text: Uniformize Font Size to 12 pt
            for r in p.findall('.//w:r', namespaces=NS):
                rPr = ensure_child(r, 'rPr')
                # Reset inline size overrides to 12 pt
                ensure_child(rPr, 'sz').attrib[f'{{{W_NS}}}val'] = '24'
                ensure_child(rPr, 'szCs').attrib[f'{{{W_NS}}}val'] = '24'

            # 6. Abnormal Bold Cleanup on Body Paragraphs
            # If paragraph text > 70 chars and predominantly bold (>70% of characters):
            runs = p.findall('.//w:r', namespaces=NS)
            if len(full_text) > 70 and runs:
                bold_chars = 0
                for r in runs:
                    rPr = r.find('w:rPr', namespaces=NS)
                    if rPr is not None and (rPr.find('w:b', namespaces=NS) is not None or rPr.find('w:bCs', namespaces=NS) is not None):
                        b = rPr.find('w:b', namespaces=NS)
                        if b is None or b.attrib.get(f'{{{W_NS}}}val', '1') not in ['0', 'false']:
                            bold_chars += len("".join(node.text for node in r.findall('.//w:t', namespaces=NS) if node.text))
                            
                if bold_chars / len(full_text) > 0.7:
                    # Predominantly bold paragraph! Check for lead-in label
                    m = lead_in_regex.match(full_text)
                    if m:
                        label_part = m.group(1).strip()
                        rest_part = m.group(2).strip()
                        
                        # Clear existing runs and rebuild with clean lead-in bold
                        for r in list(p.findall('w:r', namespaces=NS)):
                            p.remove(r)
                            
                        # Run 1: Label (Bold, 12pt)
                        r1 = etree.SubElement(p, f'{{{W_NS}}}r')
                        rPr1 = etree.SubElement(r1, f'{{{W_NS}}}rPr')
                        rf1 = etree.SubElement(rPr1, f'{{{W_NS}}}rFonts')
                        rf1.attrib[f'{{{W_NS}}}ascii'] = 'Times New Roman'
                        rf1.attrib[f'{{{W_NS}}}hAnsi'] = 'Times New Roman'
                        etree.SubElement(rPr1, f'{{{W_NS}}}b')
                        etree.SubElement(rPr1, f'{{{W_NS}}}bCs')
                        etree.SubElement(rPr1, f'{{{W_NS}}}sz').attrib[f'{{{W_NS}}}val'] = '24'
                        etree.SubElement(rPr1, f'{{{W_NS}}}szCs').attrib[f'{{{W_NS}}}val'] = '24'
                        etree.SubElement(rPr1, f'{{{W_NS}}}color').attrib[f'{{{W_NS}}}val'] = '000000'
                        t1 = etree.SubElement(r1, f'{{{W_NS}}}t')
                        t1.text = label_part + " "
                        t1.attrib['{http://www.w3.org/XML/1998/namespace}space'] = 'preserve'
                        
                        # Run 2: Narrative body (NOT Bold, 12pt)
                        r2 = etree.SubElement(p, f'{{{W_NS}}}r')
                        rPr2 = etree.SubElement(r2, f'{{{W_NS}}}rPr')
                        rf2 = etree.SubElement(rPr2, f'{{{W_NS}}}rFonts')
                        rf2.attrib[f'{{{W_NS}}}ascii'] = 'Times New Roman'
                        rf2.attrib[f'{{{W_NS}}}hAnsi'] = 'Times New Roman'
                        etree.SubElement(rPr2, f'{{{W_NS}}}sz').attrib[f'{{{W_NS}}}val'] = '24'
                        etree.SubElement(rPr2, f'{{{W_NS}}}szCs').attrib[f'{{{W_NS}}}val'] = '24'
                        etree.SubElement(rPr2, f'{{{W_NS}}}color').attrib[f'{{{W_NS}}}val'] = '000000'
                        t2 = etree.SubElement(r2, f'{{{W_NS}}}t')
                        t2.text = rest_part
                        
                        unbolded_paragraphs += 1
                    else:
                        # No lead-in: unbold all runs in this paragraph
                        for r in runs:
                            rPr = r.find('w:rPr', namespaces=NS)
                            if rPr is not None:
                                for b_tag in ['b', 'bCs']:
                                    b_elem = rPr.find(f'w:{b_tag}', namespaces=NS)
                                    if b_elem is not None:
                                        rPr.remove(b_elem)
                        unbolded_paragraphs += 1

        print(f"Cleaned abnormal bold in {unbolded_paragraphs} body paragraphs (lead-in bold applied)")
        print(f"Enforced left-alignment on {mc_aligned} multiple choice paragraphs")

        # -------------------------------------------------------------
        # 3. ENSURE ALL COLORS ARE 100% BLACK
        # -------------------------------------------------------------
        for c in doc_xml.findall('.//w:color', namespaces=NS):
            c.attrib[f'{{{W_NS}}}val'] = '000000'

        # -------------------------------------------------------------
        # 4. REPACK ARCHIVE
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
                else:
                    data = zin.read(item)
                    zout.writestr(item, data)

    if os.path.exists(dst_path):
        os.remove(dst_path)
    os.rename(tmp_dst, dst_path)

    print(f"\n==================================================")
    print(f"SUCCESS: Refined document saved to {dst_path}")
    print(f"Final size: {os.path.getsize(dst_path) / (1024*1024):.2f} MB")
    print(f"==================================================")

if __name__ == '__main__':
    src = 'documents/Tugas 4 (new).docx'
    dst = 'documents/Tugas 4 (new).docx'
    apply_refinements(src, dst)
