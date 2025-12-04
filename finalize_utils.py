import os
import shutil
from datetime import datetime
from logger import log
from db_utils import save_download_db, load_download_db
from config import OUTPUT_DIR

def check_report_ready(date_str, local_loader):

    records = local_loader(date_str)
    if not records:
        log(f"No records found locally for {date_str} — cannot finalize.")
        return False

    photos_dir = os.path.join(__import__('config').LOCAL_DIR, date_str, "photos")

    found_shift2_end = False
    for r in records:
        if r.get("type") == "end_shift" and str(r.get("shift", "")) == "2":
            photo = r.get("photo")
            if photo:
                if os.path.exists(os.path.join(photos_dir, photo)):
                    found_shift2_end = True
                    break
                else:
                    log(f"Shift2 end record references photo {photo} but file missing.")
                    return False
            else:
                log("Shift2 end record has no photo field.")
                return False

    if not found_shift2_end:
        log("Shift 2 end not found yet — not ready.")
        return False

    for r in records:
        if r.get("type") == "record_update":
            photo = r.get("photo")
            if not photo:
                log(f"Record {r.get('_filename','?')} missing photo field — not ready.")
                return False
            if not os.path.exists(os.path.join(photos_dir, photo)):
                log(f"Record {r.get('_filename','?')} references missing photo {photo} — not ready.")
                return False

    log(f"All checks passed — {date_str} is ready to finalize.")
    return True

def finalize_report(date_str, partial_docx_path=None):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    final_dir = os.path.join(OUTPUT_DIR, "final")
    os.makedirs(final_dir, exist_ok=True)

    if partial_docx_path is None:
        partial_docx_path = os.path.join(OUTPUT_DIR, f"Daily_Report_{date_str}_partial.docx")

    if not os.path.exists(partial_docx_path):
        log(f"Partial doc not found to finalize: {partial_docx_path}")
        return None

    base_name = f"Daily_Report_{date_str}_FINAL.docx"
    final_path = os.path.join(final_dir, base_name)

    if os.path.exists(final_path):
        idx = 2
        while True:
            candidate = os.path.join(final_dir, f"Daily_Report_{date_str}_FINAL_v{idx}.docx")
            if not os.path.exists(candidate):
                final_path = candidate
                break
            idx += 1

    try:
        try:
            os.replace(partial_docx_path, final_path)
            log(f"Moved partial to final: {final_path}")
        except PermissionError:
            shutil.copy2(partial_docx_path, final_path)
            log(f"Partial file locked; copied to final path: {final_path}")
    except Exception as e:
        log(f"Failed to finalize report: {e}")
        return None

    db = load_download_db()
    if date_str not in db:
        db[date_str] = {}
    db[date_str]["finalized"] = True
    db[date_str]["finalized_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_download_db(db)

    return final_path
