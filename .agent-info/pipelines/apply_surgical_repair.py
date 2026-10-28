import zipfile
import os
from lxml import etree
import docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def apply_repair(src_path, dst_path):
    print(f"Reading from baseline: {src_path}")
    with zipfile.ZipFile(src_path, 'r') as zin:
        namelist = zin.namelist()
        
        # 1. Surgical repair on word/styles.xml
        styles_xml = etree.fromstring(zin.read('word/styles.xml'))
        missing_count = 0
        for s in styles_xml.findall(f'{{{W_NS}}}style'):
            name_el = s.find(f'{{{W_NS}}}name')
            if name_el is None:
                missing_count += 1
                sid = s.attrib.get(f'{{{W_NS}}}styleId', 'Unknown')
                new_name = etree.Element(f'{{{W_NS}}}name')
                new_name.attrib[f'{{{W_NS}}}val'] = f'Table {sid}'
                s.insert(0, new_name)
        
        # Update Normal style to Times New Roman 12pt, 1.5 line, justified
        for s in styles_xml.findall(f'{{{W_NS}}}style'):
            if s.attrib.get(f'{{{W_NS}}}styleId') == 'Normal':
                pPr = s.find(f'{{{W_NS}}}pPr')
                if pPr is None:
                    pPr = etree.SubElement(s, f'{{{W_NS}}}pPr')
                spacing = pPr.find(f'{{{W_NS}}}spacing')
                if spacing is None:
                    spacing = etree.SubElement(pPr, f'{{{W_NS}}}spacing')
                spacing.attrib[f'{{{W_NS}}}after'] = '0'
                spacing.attrib[f'{{{W_NS}}}line'] = '360'
                spacing.attrib[f'{{{W_NS}}}lineRule'] = 'auto'
                
                jc = pPr.find(f'{{{W_NS}}}jc')
                if jc is None:
                    jc = etree.SubElement(pPr, f'{{{W_NS}}}jc')
                jc.attrib[f'{{{W_NS}}}val'] = 'both'
                
                rPr = s.find(f'{{{W_NS}}}rPr')
                if rPr is None:
                    rPr = etree.SubElement(s, f'{{{W_NS}}}rPr')
                rFonts = rPr.find(f'{{{W_NS}}}rFonts')
                if rFonts is None:
                    rFonts = etree.SubElement(rPr, f'{{{W_NS}}}rFonts')
                rFonts.attrib[f'{{{W_NS}}}ascii'] = 'Times New Roman'
                rFonts.attrib[f'{{{W_NS}}}hAnsi'] = 'Times New Roman'
                rFonts.attrib[f'{{{W_NS}}}cs'] = 'Times New Roman'
                
                sz = rPr.find(f'{{{W_NS}}}sz')
                if sz is None:
                    sz = etree.SubElement(rPr, f'{{{W_NS}}}sz')
                sz.attrib[f'{{{W_NS}}}val'] = '24'
                
                szCs = rPr.find(f'{{{W_NS}}}szCs')
                if szCs is None:
                    szCs = etree.SubElement(rPr, f'{{{W_NS}}}szCs')
                szCs.attrib[f'{{{W_NS}}}val'] = '24'
                
        # 2. Surgical repair on word/document.xml across all sectPr elements
        doc_xml = etree.fromstring(zin.read('word/document.xml'))
        sectPr_count = 0
        for sectPr in doc_xml.findall(f'.//{{{W_NS}}}sectPr'):
            sectPr_count += 1
            pgMar = sectPr.find(f'{{{W_NS}}}pgMar')
            if pgMar is not None:
                pgMar.attrib[f'{{{W_NS}}}top'] = '2268'
                pgMar.attrib[f'{{{W_NS}}}left'] = '2268'
                pgMar.attrib[f'{{{W_NS}}}bottom'] = '1701'
                pgMar.attrib[f'{{{W_NS}}}right'] = '1701'
            pgSz = sectPr.find(f'{{{W_NS}}}pgSz')
            if pgSz is not None:
                pgSz.attrib[f'{{{W_NS}}}w'] = '12240'
                pgSz.attrib[f'{{{W_NS}}}h'] = '15840'
                
        # 3. Clean word/settings.xml (strip embed fonts)
        settings_xml = etree.fromstring(zin.read('word/settings.xml'))
        for el in list(settings_xml):
            tag = el.tag.split('}')[-1]
            if tag in ['embedTrueTypeFonts', 'saveSubsetFonts', 'doNotEmbedSystemFonts']:
                settings_xml.remove(el)
                
        # 4. Clean word/fontTable.xml (strip embedded font references)
        fonttable_xml = etree.fromstring(zin.read('word/fontTable.xml'))
        for font in fonttable_xml.findall(f'{{{W_NS}}}font'):
            for child in list(font):
                tag = child.tag.split('}')[-1]
                if tag.startswith('embed'):
                    font.remove(child)
                    
        # Write to destination with ZIP_DEFLATED, omitting word/fonts/ and word/_rels/fontTable.xml.rels
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

    print(f"Successfully repaired -> {dst_path}")
    print(f"Patched malformed styles: {missing_count}")
    print(f"Updated section count: {sectPr_count}")
    print(f"Final file size: {os.path.getsize(dst_path) / (1024*1024):.2f} MB")

    # Run automated validation
    print("\nRunning automated validation:")
    doc = docx.Document(dst_path)
    print(f"- Paragraphs: {len(doc.paragraphs)}")
    print(f"- Tables: {len(doc.tables)}")
    for i, sec in enumerate(doc.sections):
        print(f"- Section {i}: {sec.page_width.cm:.2f} x {sec.page_height.cm:.2f} cm | Margins: Top {sec.top_margin.cm:.2f} cm, Left {sec.left_margin.cm:.2f} cm, Bottom {sec.bottom_margin.cm:.2f} cm, Right {sec.right_margin.cm:.2f} cm")

    with zipfile.ZipFile(dst_path) as z:
        st = etree.fromstring(z.read('word/styles.xml'))
        missing = sum(1 for s in st.findall(f'{{{W_NS}}}style') if s.find(f'{{{W_NS}}}name') is None)
        print(f"- Missing <w:name> tags in styles: {missing}")

if __name__ == '__main__':
    apply_repair('documents/Tugas 4_backup.docx', 'documents/Tugas 4.docx')
