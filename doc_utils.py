import os
import shutil
import uuid
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from logger import log
from config import TEMPLATE_ORIG, OUTPUT_DIR
from data_utils import find_shift_sign_photos, load_day_records_local
from image_utils import resize_image_fixed
from xml_utils import inject_images_into_docx
from xml_utils import ensure_dir
from db_utils import save_download_db, load_download_db

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
    for p in doc.paragraphs:
        if placeholder in clean_text(p.text):
            clear_and_insert(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if placeholder in clean_text(p.text):
                        clear_and_insert(p, cell)
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

def create_partial_report_with_shift_signs(date_str):
    log(f"Creating partial report for {date_str}")

    sign_map = find_shift_sign_photos(date_str)

    try:
        doc = Document(TEMPLATE_ORIG)
    except Exception as e:
        log(f"Failed to open template {TEMPLATE_ORIG}: {e}")
        return None

    human_date = datetime.strptime(date_str, "%Y-%m-%d").strftime("%d %B %Y")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    tmp_text_docx = os.path.join(OUTPUT_DIR, f"temp_text_{date_str}_{uuid.uuid4().hex}.docx")
    try:
        # replace text placeholders that python-docx can see (but not text inside textboxes)
        mapping = {
            "(date_with_month)": human_date
        }
        replace_text_placeholders(doc, mapping)
        doc.save(tmp_text_docx)
    except Exception as e:
        log(f"Failed to save temporary text document: {e}")
        return None

    placeholder_image_map = {
        "(shift_1_signin)": sign_map.get("shift_1_signin"),
        "(shift_1_signout)": sign_map.get("shift_1_signout"),
        "(shift_2_signin)": sign_map.get("shift_2_signin"),
        "(shift_2_signout)": sign_map.get("shift_2_signout"),
    }

    final_image_map = {}
    for placeholder, img_path in placeholder_image_map.items():
        if img_path:
            base, ext = os.path.splitext(img_path)
            resized_path = base + "_162.jpg"
            try:
                resize_image_fixed(img_path, resized_path, 162, 162)
                final_image_map[placeholder] = resized_path
                log(f"Resized image for {placeholder} -> {resized_path}")
            except Exception as e:
                log(f"Resize failed for {img_path}: {e}")
                final_image_map[placeholder] = img_path
        else:
            final_image_map[placeholder] = None

    placeholder_image_map = final_image_map

    final_docx = os.path.join(OUTPUT_DIR, f"Daily_Report_{date_str}_partial.docx")
    final_docx_safe = safe_save_docx(final_docx)

    mapping_for_xml = {
        "(date)": date_str
    }

    try:
        inject_images_into_docx(tmp_text_docx, final_docx_safe, placeholder_image_map, text_map=mapping_for_xml)
    except Exception as e:
        log(f"XML injection failed: {e}")
        try:
            shutil.copy2(tmp_text_docx, final_docx_safe)
        except Exception as e2:
            log(f"Failed fallback copy: {e2}")
            return None

    try:
        doc_final = Document(final_docx_safe)
        force_arial(doc_final)
        doc_final.save(final_docx_safe)
    except Exception as e:
        log(f"Post python-docx formatting failed: {e}")

    try:
        if os.path.exists(tmp_text_docx):
            os.remove(tmp_text_docx)
    except:
        pass

    log(f"Saved partial: {final_docx_safe}")
    return final_docx_safe
