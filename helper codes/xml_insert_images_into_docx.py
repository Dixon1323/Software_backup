#!/usr/bin/env python3
"""
xml_insert_images_into_docx.py

Given an input .docx and a mapping of placeholders -> image file paths,
this script injects the images directly into the DOCX XML where the placeholders are found.

Usage: edit INPUT_DOCX, OUTPUT_DOCX and PLACEHOLDERS below and run.
"""

import os
import zipfile
import shutil
import uuid
from PIL import Image
from lxml import etree

# ---------- CONFIG ----------
INPUT_DOCX = r"D:\Dixon\Automation_software\Office Software\template_formatted.docx"   # your uploaded template (from upload)
OUTPUT_DOCX = r"D:\Dixon\Automation_software\Office Software\sync\reports\template_with_images.docx"

# Map placeholder text (exact) to an image file on disk
PLACEHOLDER_TO_IMAGE = {
    "(shift_1_signin)": r"D:\Dixon\Automation_software\Office Software\sync\records\2025-11-24\photos\001.jpg",
    "(shift_1_signout)": r"D:\Dixon\Automation_software\Office Software\sync\records\2025-11-24\photos\002.jpg",
    "(shift_2_signin)": r"D:\Dixon\Automation_software\Office Software\sync\records\2025-11-24\photos\003.jpg",
    "(shift_2_signout)": r"D:\Dixon\Automation_software\Office Software\sync\records\2025-11-24\photos\004.jpg",
}

# You can change the image extension output (png/jpg)
MEDIA_EXT = ".png"
# DPI conversion: EMU per pixel multiplier
EMU_PER_PIXEL = 9525

# ---------- NAMESPACES ----------
nsmap = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

# ---------- HELPERS ----------
def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def next_media_name(media_dir):
    # choose next imgNNN.png name
    existing = [f for f in os.listdir(media_dir) if f.startswith("image") or f.startswith("img")]
    # generate unique name
    i = 1
    while True:
        name = f"image{i:03d}{MEDIA_EXT}"
        if name not in existing:
            return name
        i += 1

def add_image_file_to_media(tmpdir, image_src):
    """
    Copies/converts image_src into tmpdir/word/media/<name> and returns the filename and size in pixels.
    """
    media_dir = os.path.join(tmpdir, "word", "media")
    ensure_dir(media_dir)

    # choose a name
    fname = next_media_name(media_dir)
    out_path = os.path.join(media_dir, fname)

    # load and save as png to ensure compatibility
    img = Image.open(image_src)
    # preserve mode
    if MEDIA_EXT.lower() == ".png":
        img.save(out_path, format="PNG")
    else:
        img.save(out_path)

    width_px, height_px = img.size
    img.close()
    return fname, width_px, height_px

def ensure_rels_for_part(tmpdir, xml_rel_path):
    """
    Ensure the .rels file exists for the xml part; return its full path.
    xml_rel_path is something like 'word/_rels/document.xml.rels' or 'word/_rels/header1.xml.rels'
    """
    rels_full = os.path.join(tmpdir, xml_rel_path)
    rels_dir = os.path.dirname(rels_full)
    ensure_dir(rels_dir)
    if not os.path.exists(rels_full):
        # create minimal rels root
        root = etree.Element("Relationships", nsmap={'': 'http://schemas.openxmlformats.org/package/2006/relationships'})
        tree = etree.ElementTree(root)
        tree.write(rels_full, xml_declaration=True, encoding="utf-8")
    return rels_full

def add_image_relationship(rels_full_path, target_media_filename):
    """
    Add a Relationship entry referencing media/<target_media_filename> in rels_full_path.
    Returns the new rId string used.
    """
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(rels_full_path, parser)
    root = tree.getroot()

    # compute next available Id (rIdN)
    existing_ids = [el.get('Id') for el in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')]
    # find max numeric suffix
    maxn = 0
    for i in existing_ids:
        if i and i.startswith("rId"):
            try:
                n = int(i[3:])
                if n > maxn: maxn = n
            except:
                pass
    new_id = f"rId{maxn + 1}"

    # create new Relationship element
    RelTag = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    rel = etree.SubElement(root, RelTag)
    rel.set("Id", new_id)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    rel.set("Target", "media/" + target_media_filename)

    # write back
    tree.write(rels_full_path, xml_declaration=True, encoding="utf-8")
    return new_id

def build_drawing_xml(rel_id, cx, cy):
    """
    Return a drawing XML string that inserts the image with relationship rel_id
    cx,cy are extents in EMU (English Metric Units)
    We'll use a standard inline wp:inline snippet that Word understands.
    """
    # create a compact XML string (namespaces will be declared by parent doc when rezip)
    drawing_xml = f'''
    <w:r xmlns:w="{nsmap['w']}" xmlns:wp="{nsmap['wp']}" xmlns:a="{nsmap['a']}" xmlns:pic="{nsmap['pic']}" xmlns:r="{nsmap['r']}">
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="{cx}" cy="{cy}"/>
          <wp:docPr id="1" name="InsertedImage"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id="0" name="Picture"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="{rel_id}"/>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="{cx}" cy="{cy}"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
    '''
    # compact whitespace
    return " ".join(drawing_xml.split())

# ---------- MAIN FUNCTION ----------
def inject_images_into_docx(input_docx, output_docx, placeholder_image_map):
    if not os.path.exists(input_docx):
        raise FileNotFoundError("Input docx not found: " + str(input_docx))

    tmpdir = input_docx + "_tmp_" + uuid.uuid4().hex
    if os.path.exists(tmpdir):
        shutil.rmtree(tmpdir)
    os.makedirs(tmpdir)

    # unzip
    with zipfile.ZipFile(input_docx, 'r') as zin:
        zin.extractall(tmpdir)

    word_dir = os.path.join(tmpdir, "word")
    ensure_dir(word_dir)

    # ensure media folder exists
    media_dir = os.path.join(word_dir, "media")
    ensure_dir(media_dir)

    # track names added
    added_media = {}

    # iterate through xml parts under word/
    for root, dirs, files in os.walk(word_dir):
        for fname in files:
            if not fname.endswith(".xml"):
                continue
            xml_path = os.path.join(root, fname)
            rels_name = os.path.join(os.path.relpath(root, tmpdir), "_rels", fname + ".rels")  # e.g., word/_rels/document.xml.rels
            rels_full = os.path.join(tmpdir, rels_name) if os.path.dirname(rels_name) != "" else None

            # read raw xml text
            with open(xml_path, 'rb') as f:
                try:
                    txt = f.read().decode('utf-8')
                except UnicodeDecodeError:
                    # skip if cannot decode (unlikely)
                    continue

            modified = False

            # for each placeholder that hasn't been processed in this xml, check and replace
            for placeholder, img_path in placeholder_image_map.items():
                if placeholder not in txt:
                    continue

                if not os.path.exists(img_path):
                    print("Image missing for placeholder", placeholder, ":", img_path)
                    continue

                # add image to media folder (once)
                if img_path in added_media:
                    media_fname, w_px, h_px = added_media[img_path]
                else:
                    media_fname, w_px, h_px = add_image_file_to_media(tmpdir, img_path)
                    added_media[img_path] = (media_fname, w_px, h_px)

                # ensure rels file exists for this xml part
                # compute the correct rels path: xml_path relative to tmpdir is e.g. word/document.xml -> rels at word/_rels/document.xml.rels
                xml_rel = os.path.relpath(xml_path, tmpdir)
                rels_path = os.path.join(os.path.dirname(xml_rel), "_rels", os.path.basename(xml_rel) + ".rels")
                rels_full_path = os.path.join(tmpdir, rels_path)
                ensure_dir(os.path.dirname(rels_full_path))
                if not os.path.exists(rels_full_path):
                    # create minimal relationships root
                    root_rels = etree.Element("{http://schemas.openxmlformats.org/package/2006/relationships}Relationships")
                    tree = etree.ElementTree(root_rels)
                    tree.write(rels_full_path, xml_declaration=True, encoding="utf-8")

                # add rel and get rId
                rId = add_image_relationship(rels_full_path, media_fname)

                # compute extents in EMU
                cx = int(w_px * EMU_PER_PIXEL)
                cy = int(h_px * EMU_PER_PIXEL)

                # build drawing XML referencing rId
                drawing_snippet = build_drawing_xml(rId, cx, cy)

                # replace EXACT placeholder string with drawing XML (textual replacement)
                txt = txt.replace(placeholder, drawing_snippet)
                modified = True

            if modified:
                # write back modified xml
                with open(xml_path, 'wb') as f:
                    f.write(txt.encode('utf-8'))

    # rezip to output
    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
        for foldername, subfolders, filenames in os.walk(tmpdir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, tmpdir)
                zout.write(filepath, arcname)

    # cleanup
    shutil.rmtree(tmpdir)
    return output_docx

# ---------- RUN ----------
if __name__ == "__main__":
    # quick guard
    if not PLACEHOLDER_TO_IMAGE:
        print("Set PLACEHOLDER_TO_IMAGE mapping in the script first.")
    else:
        out = inject_images_into_docx(INPUT_DOCX, OUTPUT_DOCX, PLACEHOLDER_TO_IMAGE)
        print("Saved:", out)
