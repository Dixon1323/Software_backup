#!/usr/bin/env python3
# fix_docx_drawings.py
# Replaces <w:drawing> elements in word/document.xml by normal <w:r><w:t>text</w:t></w:r>
# Usage:
#   python fix_docx_drawings.py input.docx output.docx

import sys, zipfile, shutil, os
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'v': 'urn:schemas-microsoft-com:vml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
}

def extract_text_from_drawing(drawing_elem):
    texts = drawing_elem.xpath('.//a:t', namespaces=NS)
    pieces = [t.text for t in texts if t is not None]
    return ''.join(pieces).strip()

def fix_document_xml(data_bytes):
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    tree = etree.fromstring(data_bytes, parser=parser)
    drawings = tree.xpath('.//w:drawing', namespaces=NS)
    replaced = 0
    for dr in drawings:
        text = extract_text_from_drawing(dr)
        if not text:
            continue
        # create w:r/w:t
        r = etree.Element('{%s}r' % NS['w'])
        t = etree.SubElement(r, '{%s}t' % NS['w'])
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        parent = dr.getparent()
        if parent is not None:
            parent.replace(dr, r)
            replaced += 1
    return etree.tostring(tree, xml_declaration=True, encoding='utf-8', standalone=False), replaced

def fix_docx(input_docx, output_docx):
    if not os.path.exists(input_docx):
        raise FileNotFoundError(input_docx)
    with zipfile.ZipFile(input_docx, 'r') as zin:
        namelist = zin.namelist()
        if 'word/document.xml' not in namelist:
            raise RuntimeError("word/document.xml not found in docx")
        doc_xml = zin.read('word/document.xml')
        new_doc_xml, replaced = fix_document_xml(doc_xml)
        # build output
        with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in namelist:
                if item == 'word/document.xml':
                    zout.writestr(item, new_doc_xml)
                else:
                    zout.writestr(item, zin.read(item))
    return replaced

if __name__ == "__main__":

    inp = r"D:\Dixon\Automation_software\Office Software\template_formatted.docx"
    outp = r"D:\Dixon\Automation_software\Office Software\template_formatted_new.docx"
    count = fix_docx(inp, outp)
    print(f"Finished. Replaced ~{count} drawing elements. Saved: {outp}")
