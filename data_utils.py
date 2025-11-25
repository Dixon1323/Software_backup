import os
import json
from config import LOCAL_DIR
from logger import log

def load_day_records_local(date_str):
    data_folder = os.path.join(LOCAL_DIR, date_str, "data")
    records = []
    if not os.path.exists(data_folder):
        return records
    for fname in sorted(os.listdir(data_folder)):
        if not fname.lower().endswith(".json"):
            continue
        try:
            with open(os.path.join(data_folder, fname), "r", encoding="utf-8") as f:
                rec = json.load(f)
                rec["_filename"] = fname
                records.append(rec)
        except Exception as e:
            log(f"Could not read {fname}: {e}")
    return records

def find_shift_sign_photos(date_str):
    result = {
        "shift_1_signin": None,
        "shift_1_signout": None,
        "shift_2_signin": None,
        "shift_2_signout": None
    }
    records = load_day_records_local(date_str)
    photo_dir = os.path.join(__import__('config').LOCAL_DIR, date_str, "photos")
    for r in records:
        r_type = r.get("type")
        shift = str(r.get("shift", ""))
        photo = r.get("photo")
        if not photo:
            continue
        full_path = os.path.join(photo_dir, photo)
        if not os.path.exists(full_path):
            continue
        if r_type == "start_shift" and shift == "1" and result["shift_1_signin"] is None:
            result["shift_1_signin"] = full_path
        if r_type == "end_shift" and shift == "1" and result["shift_1_signout"] is None:
            result["shift_1_signout"] = full_path
        if r_type == "start_shift" and shift == "2" and result["shift_2_signin"] is None:
            result["shift_2_signin"] = full_path
        if r_type == "end_shift" and shift == "2" and result["shift_2_signout"] is None:
            result["shift_2_signout"] = full_path
    return result
