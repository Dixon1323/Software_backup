# doc_utils.py
import os
import shutil
import uuid
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

# local imports (project modules)
from logger import log
from config import TEMPLATE_ORIG, OUTPUT_DIR, LOCAL_DIR
from data_utils import find_shift_sign_photos, load_day_records_local
from image_utils import resize_image_fixed
from xml_utils import inject_images_into_docx, ensure_dir
from db_utils import save_download_db, load_download_db

# --- helper functions (kept minimal and robust) ---

def inline_replace_paragraph(paragraph, target, replacement):
    if target not in paragraph.text:
        return False
    new_text = paragraph.text.replace(target, str(replacement))
    for r in paragraph.runs:
        r.text = ""
    paragraph.add_run(new_text)
    return True

def replace_text_placeholders(doc, mapping):
    # paragraphs
    for p in doc.paragraphs:
        for key, val in mapping.items():
            if key in p.text:
                inline_replace_paragraph(p, key, val)
    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in mapping.items():
                        if key in p.text:
                            inline_replace_paragraph(p, key, val)
    # headers/footers
    try:
        for section in doc.sections:
            for p in section.header.paragraphs:
                for key, val in mapping.items():
                    if key in p.text:
                        inline_replace_paragraph(p, key, val)
            for p in section.footer.paragraphs:
                for key, val in mapping.items():
                    if key in p.text:
                        inline_replace_paragraph(p, key, val)
    except Exception:
        pass

def insert_image_at_placeholder(doc, placeholder, image_path, width_inches=2.8):
    """
    Fallback insertion using python-docx when placeholder is plain text (not inside drawing/textbox).
    Returns number of insertions done.
    """
    width = Inches(width_inches)
    inserted = 0

    def clean_text(s):
        if s is None:
            return ""
        return s.replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")

    def clear_and_insert(paragraph, cell=None):
        nonlocal inserted
        try:
            if cell:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = ""
                p = cell.paragraphs[0]
            else:
                for r in paragraph.runs:
                    r.text = ""
                p = paragraph
            run = p.add_run()
            run.add_picture(image_path, width=width)
            inserted += 1
        except Exception as e:
            log(f"Image insert failed at {placeholder}: {e}")

    # paragraphs
    for p in doc.paragraphs:
        if placeholder in clean_text(p.text):
            clear_and_insert(p)

    # table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if placeholder in clean_text(p.text):
                        clear_and_insert(p, cell)

    # headers / footers
    try:
        for section in doc.sections:
            for p in section.header.paragraphs:
                if placeholder in clean_text(p.text):
                    clear_and_insert(p)
            for p in section.footer.paragraphs:
                if placeholder in clean_text(p.text):
                    clear_and_insert(p)
    except Exception:
        pass

    return inserted

def force_arial(doc, size_pt=11):
    for p in doc.paragraphs:
        for run in p.runs:
            try:
                run.font.name = "Arial"
                run.font.size = Pt(size_pt)
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            except Exception:
                pass
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        try:
                            run.font.name = "Arial"
                            run.font.size = Pt(size_pt)
                            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
                        except Exception:
                            pass

def safe_save_docx(base_path):
    """
    If file exists and is locked by Word, returns a new versioned filename
    (e.g., base_v2.docx) so we don't raise PermissionError when writing.
    """
    if not os.path.exists(base_path):
        return base_path
    try:
        os.rename(base_path, base_path)
        return base_path
    except PermissionError:
        folder = os.path.dirname(base_path)
        name = os.path.basename(base_path)
        base, ext = os.path.splitext(name)
        version = 2
        while True:
            new_name = f"{base}_v{version}{ext}"
            new_path = os.path.join(folder, new_name)
            if not os.path.exists(new_path):
                return new_path
            version += 1

# --- new helper: scan photos folder and map (pic_<num>) placeholders to resized image paths ---

def build_pic_placeholders_map(date_str, desired_w=162, desired_h=162):
    """
    Scans LOCAL_DIR/<date>/photos and returns dict mapping '(pic_<cage>)' -> resized_image_path
    Example: '002.jpg' -> placeholder '(pic_002)' and '(pic_2)' (both)
    """
    photos_dir = os.path.join(LOCAL_DIR, date_str, "photos")
    mapping = {}
    if not os.path.exists(photos_dir):
        return mapping

    for fname in sorted(os.listdir(photos_dir)):
        if not fname.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        src = os.path.join(photos_dir, fname)
        # derive digits from filename
        name_root = os.path.splitext(fname)[0]
        digits = ''.join([c for c in name_root if c.isdigit()])
        if not digits:
            digits = name_root

        placeholder1 = f"(pic_{digits})"
        placeholder2 = f"(pic_{int(digits)})" if digits.isdigit() else placeholder1

        base, ext = os.path.splitext(src)
        resized_path = base + "_162.jpg"
        try:
            resize_image_fixed(src, resized_path, desired_w, desired_h)
            mapping[placeholder1] = resized_path
            if placeholder2 != placeholder1:
                mapping[placeholder2] = resized_path
            log(f"Prepared resized cage photo {fname} -> {os.path.basename(resized_path)}")
        except Exception as e:
            log(f"Resize failed for {src}: {e}")
            mapping[placeholder1] = src
            if placeholder2 != placeholder1:
                mapping[placeholder2] = src

    return mapping


# --- PLACE -> list of cage numbers ---
PLACES_CAGES = {
    "Southern Promenade": [458,459,460,461,462,463,464,465,466,467,468,469,470,471,472,473,474],
    "Eastern Promenade": [475,476,477,478,526,527,528,530,531],
    "U-shape East & West Wing": [484,485,486,487,488,489,490,491,492,501,502,505],
    "Marina Carpark 2A": [506,507,508,510],
    "Marina Carpark 2B": [493,494,495],
    "Northern Promenade": [512,513],
    "QD Complex (External)": [496,497,498,499,500,504,509],
    "Crescent Park 01": [523,524,525,540,541,542,543,544,545,546,547],
    "Crescent Park 02": [479,480,481,482,483],
    "Crescent Park 03": [514,516,517,518,519,520,521,522],
    "Crescent Park 04": [503,511,532,533,534,535,536,537,538,539],
    "Crescent Park 05": [549,550,551,552,553,554,555,556,557,558,559,560,561,562,563,564],
    "Al Khuzama Zone -2": [604,605,606,607,608,609,610,611,612,613,614,615,617,618,619,620],
    "Al Khuzama Zone -1": [588,589,590,591,592,593,594,595,596,597,598,599,600,601,602,603],
    "Al Nafel Park": [577,578,579,580,581],
    "QETAIFAN ZONE 1": [569,570,574,575],
    "QETAIFAN ZONE 2": [567,572,573,576],
    "QETAIFAN ZONE 3": [565,566,568],
    "Qetaifan North Park": [623,624,625,626,630,631,632,633,634],
    "Road A1 - Al Khuzama": [621,622,627,628,629],
    "Seef Lusail North": [635,636,637,638,639,640,641,642],
    # ensure no duplicates above
}

# placeholder defs for totals naming conventions (shift-2 placeholders shown here; shift-1 analogous)
PLACE_TOTAL_PLACEHOLDERS = {
    "Southern Promenade": ("(2sp_total)", "(2spm)", "(2spl)"),
    "Eastern Promenade": ("(2ept)", "(2epm)", "(2epl)"),
    "U-shape East & West Wing": ("(2uset)", "(2usem)", "(2usel)"),
    "Northern Promenade": ("(2npt)", "(2npm)", "(2npl)"),
    "QD Complex (External)": ("(2qdcet)", "(2qdcetm)", "(2qdcetl)"),
    "Marina Carpark 2B": ("(2mc2bt)", "(2mc2bm)", "(2mc2bl)"),
    "Marina Carpark 2A": ("(2mc2at)", "(2mc2am)", "(2mc2al)"),
    "Crescent Park 01": ("(2cp1t)", "(2cp1m)", "(2cp1l)"),
    "Crescent Park 02": ("(2cp2t)", "(2cp2m)", "(2cp2l)"),
    "Crescent Park 03": ("(2cp3t)", "(2cp3m)", "(2cp3l)"),
    "Crescent Park 04": ("(2cp4t)", "(2cp4m)", "(2cp4l)"),
    "Crescent Park 05": ("(2cp5t)", "(2cp5m)", "(2cp5l)"),
    "QETAIFAN ZONE 1": ("(2qz1t)", "(2qz1m)", "(2qz1l)"),
    "QETAIFAN ZONE 2": ("(2qz2t)", "(2qz2m)", "(2qz2l)"),
    "QETAIFAN ZONE 3": ("(2qz3t)", "(2qz3m)", "(2qz3l)"),
    "Al Nafel Park": ("(2anpt)", "(2anpm)", "(2anpl)"),  # fixed typo: (2anpm)/(2anpl)
    "Al Khuzama Zone -2": ("(2akz2t)", "(2akz2m)", "(2akz2l)"),
    "Al Khuzama Zone -1": ("(2akz1t)", "(2akz1m)", "(2akz1l)"),
    "Road A1 - Al Khuzama": ("(2raakt)", "(2raakm)", "(2raakl)"),  # fixed typo: (2raakm)
    "Qetaifan North Park": ("(2qnpt)", "(2qnpm)", "(2qnpl)"),
    "Seef Lusail North": ("(2slnt)", "(2slnm)", "(2slnl)"),
}

# helper that returns the placeholder name for a cage in a given shift (1 or 2)
def cage_placeholder(shift, cage_number):
    return f"({shift}c{cage_number})"

# helper for picture placeholder
def pic_placeholder(cage_number):
    return f"(pic_{cage_number})"

# -------------------------
# process_record_updates:
# Reads local JSONs and constructs:
#  - text_map: placeholders -> string values (per-cage and totals)
#  - pic_map: (pic_N) -> resized image path
# -------------------------
def process_record_updates(date_str, local_dir, resize_fn, logger):
    text_map = {}
    pic_map = {}

    records = []
    data_folder = os.path.join(local_dir, date_str, "data")
    photos_folder = os.path.join(local_dir, date_str, "photos")

    # load JSONs
    if os.path.exists(data_folder):
        for fname in sorted(os.listdir(data_folder)):
            if not fname.lower().endswith(".json"):
                continue
            try:
                with open(os.path.join(data_folder, fname), "r", encoding="utf-8") as f:
                    rec = json.load(f)
                    rec["_filename"] = fname
                    records.append(rec)
            except Exception as e:
                logger(f"process_record_updates: failed reading {fname}: {e}")

    # init per-place aggregates
    place_agg = {place: {'myna': 0, 'local': 0, 'total': 0} for place in PLACES_CAGES.keys()}

    # Initialize all cage placeholders to "0" so they exist
    for place, cages in PLACES_CAGES.items():
        for c in cages:
            text_map[cage_placeholder(1, c)] = "0"
            text_map[cage_placeholder(2, c)] = "0"

    # process record updates
    for r in records:
        if r.get("type") != "record_update":
            continue
        shift = str(r.get("shift", "1")).strip()
        try:
            cage_no = int(r.get("cage_number"))
        except Exception:
            logger(f"process_record_updates: invalid cage_number in record {r.get('_filename','?')}")
            continue

        # parse counts
        myna = 0
        local = 0
        try:
            myna = int(str(r.get("myna_captured") or "0"))
        except Exception:
            myna = 0
        try:
            local = int(r.get("local_released") or 0)
        except Exception:
            local = 0

        total_for_cell = myna + local
        ph = cage_placeholder(shift, cage_no)
        text_map[ph] = str(total_for_cell)

        # zero the opposite shift
        opp_shift = "1" if shift == "2" else "2"
        opp_ph = cage_placeholder(opp_shift, cage_no)
        text_map[opp_ph] = "0"

        # update place aggregates
        for place, cages in PLACES_CAGES.items():
            if cage_no in cages:
                place_agg[place]['myna'] += myna
                place_agg[place]['local'] += local
                place_agg[place]['total'] += total_for_cell
                break

        # picture mapping
        photo = r.get("photo")
        if photo:
            src = os.path.join(photos_folder, photo)
            if os.path.exists(src):
                base, ext = os.path.splitext(src)
                dst = base + "_162.jpg"
                try:
                    resize_fn(src, dst, 162, 162)
                    pic_map[pic_placeholder(cage_no)] = dst
                except Exception as e:
                    logger(f"process_record_updates: resize failed for {src}: {e}")
                    pic_map[pic_placeholder(cage_no)] = src
            else:
                logger(f"process_record_updates: photo file for {photo} not found")

    # totals placeholders
    for place, agg in place_agg.items():
        tpl = PLACE_TOTAL_PLACEHOLDERS.get(place)
        if not tpl:
            continue
        total_ph, myna_ph, local_ph = tpl
        text_map[total_ph] = str(agg['total'])
        text_map[myna_ph] = str(agg['myna'])
        text_map[local_ph] = str(agg['local'])

    logger(f"process_record_updates: generated {len(text_map)} text placeholders and {len(pic_map)} pic placeholders")
    return text_map, pic_map

# -------------------------
# create_partial_report_with_shift_signs: orchestrates text replace + XML injection
# -------------------------
def create_partial_report_with_shift_signs(date_str):
    log(f"Creating partial report for {date_str}")

    # find sign photos
    sign_map = find_shift_sign_photos(date_str)

    # Load original template
    try:
        doc = Document(TEMPLATE_ORIG)
    except Exception as e:
        log(f"Failed to open template {TEMPLATE_ORIG}: {e}")
        return None

    # Date strings
    human_date = datetime.strptime(date_str, "%Y-%m-%d").strftime("%d %B %Y")

    # Save a temporary copy where python-docx can replace visible text (not textboxes)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    tmp_text_docx = os.path.join(OUTPUT_DIR, f"temp_text_{date_str}_{uuid.uuid4().hex}.docx")

    try:
        mapping_visible = {"(date_with_month)": human_date}
        replace_text_placeholders(doc, mapping_visible)
        doc.save(tmp_text_docx)
    except Exception as e:
        log(f"Failed to save temporary text document: {e}")
        return None

    # Build placeholders -> images map for signs first
    placeholder_image_map = {
        "(shift_1_signin)": None,
        "(shift_1_signout)": None,
        "(shift_2_signin)": None,
        "(shift_2_signout)": None
    }

    # resize sign images to 162x162 if present
    for k in ["shift_1_signin", "shift_1_signout", "shift_2_signin", "shift_2_signout"]:
        src = sign_map.get(k)
        if src and os.path.exists(src):
            base, ext = os.path.splitext(src)
            resized = base + "_162.jpg"
            try:
                resize_image_fixed(src, resized, 162, 162)
                placeholder_image_map[f"({k})"] = resized
                log(f"Resized sign {k}: {resized}")
            except Exception as e:
                log(f"Sign resize failed for {src}: {e}")
                placeholder_image_map[f"({k})"] = src
        else:
            placeholder_image_map[f"({k})"] = None

    # 2) Cage photos: scan photos folder for (pic_###) placeholders
    pic_map = build_pic_placeholders_map(date_str, desired_w=162, desired_h=162)

    # 3) process record updates -> generate numeric text map and per-cage pic_map (from record photos)
    text_map_updates, pic_map_from_records = process_record_updates(date_str, LOCAL_DIR, resize_image_fixed, log)

    # Merge pic maps (record-specific pics should override generic scanned ones)
    pic_map.update(pic_map_from_records)

    # Merge sign+pic into placeholder_image_map
    placeholder_image_map.update(pic_map)

    # Text replacements to perform at XML level (handles textboxes)
    # include date (in textbox), and **all numeric placeholders + totals** from process_record_updates
    mapping_for_xml = {
        "(date)": date_str,
        "(date_with_month)": human_date
    }
    # merge numeric placeholders and totals
    mapping_for_xml.update(text_map_updates)

    # Final doc path (safe)
    final_docx = os.path.join(OUTPUT_DIR, f"Daily_Report_{date_str}_partial.docx")
    final_docx_safe = safe_save_docx(final_docx)

    # Perform XML-level injection (images + text_map_for_xml)
    try:
        inject_images_into_docx(tmp_text_docx, final_docx_safe, placeholder_image_map, text_map=mapping_for_xml)
    except Exception as e:
        log(f"XML injection failed: {e}")
        # fallback copy
        try:
            shutil.copy2(tmp_text_docx, final_docx_safe)
        except Exception as e2:
            log(f"Failed fallback copy: {e2}")
            return None

    # Final python-docx pass to force Arial and do any plain-text picture insertions fallback
    try:
        doc_final = Document(final_docx_safe)
        # fallback: if some pic placeholders weren't handled at XML level and are plain text, try python-docx insertion
        for placeholder, img_path in placeholder_image_map.items():
            if img_path and "(" in placeholder:
                inserted = insert_image_at_placeholder(doc_final, placeholder, img_path, width_inches=1.8)
                if inserted:
                    log(f"Fallback inserted {placeholder} via python-docx (count={inserted})")
        force_arial(doc_final)
        doc_final.save(final_docx_safe)
    except Exception as e:
        log(f"Post python-docx formatting failed: {e}")

    # cleanup temp file
    try:
        if os.path.exists(tmp_text_docx):
            os.remove(tmp_text_docx)
    except Exception:
        pass

    log(f"Saved partial: {final_docx_safe}")
    return final_docx_safe
