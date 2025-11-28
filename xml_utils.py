import os
import zipfile
import shutil
import uuid
import re
from lxml import etree
from PIL import Image
from logger import log
from config import MEDIA_EXT, EMU_PER_PIXEL


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}


def replace_placeholder_across_wt_nodes(xml_bytes, placeholder, drawing_snippet):
    """
    Try to replace placeholder that may be split across multiple <w:t> nodes.
    xml_bytes: bytes content of the XML part
    placeholder: e.g. "(1c588)" (string)
    drawing_snippet: XML string representing a <w:r> element (as in your build_drawing_xml output)
    Returns: (new_xml_bytes, replaced_count) where replaced_count is 0 or 1
    """
    try:
        parser = etree.XMLParser(ns_clean=True, recover=True, remove_blank_text=False)
        root = etree.fromstring(xml_bytes, parser=parser)
    except Exception:
        return xml_bytes, 0

    # find all text nodes in document order
    text_nodes = root.findall('.//' + W_NS + 't')
    if not text_nodes:
        return xml_bytes, 0

    # Build list of texts and cumulative lengths
    texts = [ (t, (t.text or "")) for t in text_nodes ]
    concat = "".join([t for (_, t) in texts])
    idx = concat.find(placeholder)
    if idx == -1:
        return xml_bytes, 0

    # determine which nodes cover that range
    start = 0
    s_node = None
    e_node = None
    cur = 0
    for i, (_, txt) in enumerate(texts):
        cur_len = len(txt)
        if cur + cur_len > idx and s_node is None:
            s_idx = i
            s_node = text_nodes[i]
            s_offset = idx - cur
        if cur + cur_len >= idx + len(placeholder):
            e_idx = i
            e_node = text_nodes[i]
            e_offset = (idx + len(placeholder)) - cur
            break
        cur += cur_len

    if s_node is None or e_node is None:
        return xml_bytes, 0

    # find the <w:r> parent for the start node
    start_run = s_node.getparent()
    while start_run is not None and start_run.tag != W_NS + 'r':
        start_run = start_run.getparent()
    if start_run is None:
        return xml_bytes, 0

    # Remove/clear text content across involved nodes and remove runs between start and end
    # We'll replace the start_run with the drawing snippet element
    # Parse drawing_snippet into element(s)
    try:
        drawing_el = etree.fromstring(drawing_snippet.encode('utf-8'))
    except Exception:
        # drawing_snippet might not be a full-root fragment; wrap it
        drawing_el = etree.fromstring(f"<root>{drawing_snippet}</root>".encode('utf-8'))

        # if wrapped, get the first child as the <w:r> element
        children = list(drawing_el)
        if children:
            drawing_el = children[0]
        else:
            return xml_bytes, 0

    # find all run parents between s_idx and e_idx inclusive and remove them (we'll insert drawing at start position)
    run_nodes = []
    for i in range(s_idx, e_idx + 1):
        tn = text_nodes[i]
        r = tn.getparent()
        # sometimes w:t inside other wrappers, ensure r is <w:r>
        while r is not None and r.tag != W_NS + 'r':
            r = r.getparent()
        if r is not None:
            run_nodes.append(r)

    parent_of_runs = run_nodes[0].getparent() if run_nodes else start_run.getparent()
    # insert drawing element at index of first run
    insert_index = list(parent_of_runs).index(run_nodes[0]) if run_nodes else list(parent_of_runs).index(start_run)
    # At times drawing_el may have namespaces; import it into this tree
    parent_of_runs.insert(insert_index, drawing_el)

    # remove original runs
    for r in run_nodes:
        try:
            parent_of_runs.remove(r)
        except Exception:
            # best-effort
            pass

    # write back
    new_bytes = etree.tostring(root, xml_declaration=True, encoding='utf-8')
    return new_bytes, 1


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def next_media_name(media_dir):
    existing = set(os.listdir(media_dir)) if os.path.exists(media_dir) else set()
    i = 1
    while True:
        name = f"image{i:03d}{MEDIA_EXT}"
        if name not in existing:
            return name
        i += 1

def add_image_file_to_media(tmpdir, image_src):
    media_dir = os.path.join(tmpdir, "word", "media")
    ensure_dir(media_dir)
    fname = next_media_name(media_dir)
    out_path = os.path.join(media_dir, fname)
    img = Image.open(image_src)
    try:
        if MEDIA_EXT.lower() == ".png":
            img.save(out_path, format="PNG")
        else:
            img.save(out_path)
        w_px, h_px = img.size
        img.close()
        return fname, w_px, h_px
    except Exception as e:
        img.close()
        raise

def ensure_rels_file(rels_full):
    if not os.path.exists(rels_full):
        root = etree.Element("{http://schemas.openxmlformats.org/package/2006/relationships}Relationships")
        tree = etree.ElementTree(root)
        tree.write(rels_full, xml_declaration=True, encoding="utf-8")

def add_image_relationship(rels_full_path, target_media_filename):
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(rels_full_path, parser)
    root = tree.getroot()
    existing_ids = [el.get('Id') for el in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')]
    maxn = 0
    for eid in existing_ids:
        if eid and eid.startswith("rId"):
            try:
                n = int(eid[3:])
                if n > maxn:
                    maxn = n
            except Exception:
                pass
    new_id = f"rId{maxn + 1}"
    RelTag = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    rel = etree.SubElement(root, RelTag)
    rel.set("Id", new_id)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    rel.set("Target", "media/" + target_media_filename)
    tree.write(rels_full_path, xml_declaration=True, encoding="utf-8")
    return new_id

def build_drawing_xml(rel_id, cx, cy):
    drawing_xml = f'''
    <w:r xmlns:w="{NSMAP['w']}" xmlns:wp="{NSMAP['wp']}" xmlns:a="{NSMAP['a']}" xmlns:pic="{NSMAP['pic']}" xmlns:r="{NSMAP['r']}">
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
    return " ".join(drawing_xml.split())

def inject_images_into_docx(input_docx, output_docx, placeholder_image_map, text_map=None):
    if not os.path.exists(input_docx):
        raise FileNotFoundError("Input docx not found: " + str(input_docx))

    tmpdir = input_docx + "_tmp_" + uuid.uuid4().hex
    if os.path.exists(tmpdir):
        shutil.rmtree(tmpdir)
    os.makedirs(tmpdir)

    with zipfile.ZipFile(input_docx, 'r') as zin:
        zin.extractall(tmpdir)

    word_dir = os.path.join(tmpdir, "word")
    ensure_dir(word_dir)
    media_dir = os.path.join(word_dir, "media")
    ensure_dir(media_dir)

    added_media = {}

    for root, dirs, files in os.walk(word_dir):
        for fname in files:
            if not fname.endswith(".xml"):
                continue
            xml_path = os.path.join(root, fname)
            try:
                with open(xml_path, 'rb') as f:
                    txt = f.read().decode('utf-8')
            except Exception:
                continue

            modified = False

            if text_map:
                for tkey, tval in text_map.items():
                    if tkey in txt:
                        txt = txt.replace(tkey, str(tval))
                        modified = True

            for placeholder, img_path in placeholder_image_map.items():
                if not placeholder in txt:
                    continue
                if not img_path or not os.path.exists(img_path):
                    log(f"Image missing for placeholder {placeholder}: {img_path}")
                    continue

                if img_path in added_media:
                    media_fname, w_px, h_px = added_media[img_path]
                else:
                    try:
                        media_fname, w_px, h_px = add_image_file_to_media(tmpdir, img_path)
                    except Exception as e:
                        log(f"Failed adding image to media for {img_path}: {e}")
                        continue
                    added_media[img_path] = (media_fname, w_px, h_px)

                xml_rel = os.path.relpath(xml_path, tmpdir)
                rels_path = os.path.join(os.path.dirname(xml_rel), "_rels", os.path.basename(xml_rel) + ".rels")
                rels_full = os.path.join(tmpdir, rels_path)
                ensure_dir(os.path.dirname(rels_full))
                if not os.path.exists(rels_full):
                    root_rels = etree.Element("{http://schemas.openxmlformats.org/package/2006/relationships}Relationships")
                    tree = etree.ElementTree(root_rels)
                    tree.write(rels_full, xml_declaration=True, encoding="utf-8")

                try:
                    rId = add_image_relationship(rels_full, media_fname)
                except Exception as e:
                    log(f"Failed to add relationship for media {media_fname}: {e}")
                    continue

                cx = int(w_px * EMU_PER_PIXEL)
                cy = int(h_px * EMU_PER_PIXEL)

                drawing_snippet = build_drawing_xml(rId, cx, cy)

                # try fast replace first
                if placeholder in txt:
                    new_txt = txt.replace(placeholder, drawing_snippet)
                    if new_txt != txt:
                        txt = new_txt
                        modified = True
                    else:
                        # placeholder may be split across <w:t> nodes — attempt robust node-level replace
                        try:
                            new_bytes, replaced = replace_placeholder_across_wt_nodes(txt.encode('utf-8'), placeholder, drawing_snippet)
                            if replaced:
                                txt = new_bytes.decode('utf-8')
                                modified = True
                                log(f"Performed node-level replacement for placeholder {placeholder} in {xml_path}")
                        except Exception as e:
                            log(f"Node-level replacement error for {placeholder} in {xml_path}: {e}")

                modified = True

            if modified:
                try:
                    with open(xml_path, 'wb') as f:
                        f.write(txt.encode('utf-8'))
                except Exception as e:
                    log(f"Failed to write xml part {xml_path}: {e}")

    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
        for foldername, subfolders, filenames in os.walk(tmpdir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, tmpdir)
                zout.write(filepath, arcname)

    shutil.rmtree(tmpdir)
    return output_docx
