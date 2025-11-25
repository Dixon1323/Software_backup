import os
from logger import log
from db_utils import load_download_db, save_download_db
from http_utils import safe_request, download_file, delete_from_server
from config import BASE_URL, LOCAL_DIR
from doc_utils import create_partial_report_with_shift_signs
from data_utils import load_day_records_local
from finalize_utils import check_report_ready, finalize_report

download_db = load_download_db()

def sync_day(day):
    global download_db
    new_data = False
    new_photos = False

    res = safe_request(BASE_URL + f"{day}/list")
    if res is None:
        log(f"Could not fetch file list for {day}")
        return

    files = res.json()
    server_data = files.get("data", [])
    server_photos = files.get("photos", [])

    if day not in download_db:
        download_db[day] = {"data": [], "photos": []}

    known_data = set(download_db[day]["data"])
    known_photos = set(download_db[day]["photos"])

    data_dir = os.path.join(LOCAL_DIR, day, "data")
    photos_dir = os.path.join(LOCAL_DIR, day, "photos")
    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(photos_dir, exist_ok=True)

    for f in server_data:
        if f not in known_data:
            if download_file(BASE_URL + f"{day}/data/{f}", os.path.join(data_dir, f)):
                download_db[day]["data"].append(f)
                save_download_db(download_db)
                new_data = True

    for f in server_photos:
        if f not in known_photos:
            if download_file(BASE_URL + f"{day}/photos/{f}", os.path.join(photos_dir, f)):
                download_db[day]["photos"].append(f)
                save_download_db(download_db)
                new_photos = True

    if new_data or new_photos:
        partial_path = create_partial_report_with_shift_signs(day)
        if partial_path is None:
            log("Partial report creation failed.")
        else:
            if download_db.get(day, {}).get("finalized"):
                log(f"{day} already finalized — skipping finalization.")
            else:
                if check_report_ready(day, load_day_records_local):
                    final_path = finalize_report(day, partial_docx_path=partial_path)
                    if final_path:
                        log(f"Report finalized: {final_path}")
                    else:
                        log("Finalization attempt failed.")
    else:
        log("No new files – report unchanged.")

    if len(server_data) + len(server_photos) >= 10:
        log(f"Reached limit, deleting server files for {day}")
        delete_from_server(day)
