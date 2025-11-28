# sync_day.py (patched)
import os
import json
from logger import log
from db_utils import load_download_db, save_download_db
from http_utils import safe_request, download_file, delete_from_server
from config import BASE_URL, LOCAL_DIR, OUTPUT_DIR
from doc_utils import create_partial_report_with_shift_signs
from data_utils import load_day_records_local
from finalize_utils import check_report_ready, finalize_report
from datetime import datetime

download_db = load_download_db()

# ---------------------------
# Place -> cage arrays (from your spec)
# ---------------------------
PLACE_CAGE_MAP = {
    "Al Khuzama Zone -1": [588,589,590,591,592,593,594,595,596,597,598,599,600,601,602,603],
    "Al Khuzama Zone -2": [604,605,606,607,608,609,610,611,612,613,614,615,617,618,619,620],
    "Al Nafel Park": [577,578,579,580,581],
    "Crescent Park 01": [523,524,525,540,541,542,543,544,545,546,547],
    "Crescent Park 02": [479,480,481,482,483],
    "Crescent Park 03": [514,516,517,518,519,520,521,522],
    "Crescent Park 04": [503,511,532,533,534,535,536,537,538,539],
    "Crescent Park 05": [549,550,551,552,553,554,555,556,557,558,559,560,561,562,563,564],
    "Eastern Promenade": [475,476,477,478,526,527,528,530,531],
    "Marina Carpark 2A": [506,507,508,510],
    "Marina Carpark 2B": [493,494,495],
    "Northern Promenade": [512,513],
    "QD Complex (External)": [496,497,498,499,500,504,509],
    "QETAIFAN ZONE 1": [569,570,574,575],
    "QETAIFAN ZONE 2": [567,572,573,576],
    "QETAIFAN ZONE 3": [565,566,568],
    "Qetaifan North Park": [623,624,625,626,630,631,632,633,634],
    "Road A1 - Al Khuzama": [621,622,627,628,629],
    "Seef Lusail North": [635,636,637,638,639,640,641,642],
    "Southern Promenade": [458,459,460,461,462,463,464,465,466,467,468,469,470,471,472,473,474],
    "U-shape East & West Wing": [484,485,486,487,488,489,490,491,492,501,502,505]
}

# short codes used in placeholders: these follow your naming in the template
PLACE_CODE_MAP = {
    "Southern Promenade": "sp",
    "Eastern Promenade": "ep",
    "U-shape East & West Wing": "use",
    "Northern Promenade": "np",
    "QD Complex (External)": "qdce",
    "Marina Carpark 2B": "mc2b",
    "Marina Carpark 2A": "mc2a",
    "Crescent Park 01": "cp1",
    "Crescent Park 02": "cp2",
    "Crescent Park 03": "cp3",
    "Crescent Park 04": "cp4",
    "Crescent Park 05": "cp5",
    "QETAIFAN ZONE 1": "qz1",
    "QETAIFAN ZONE 2": "qz2",
    "QETAIFAN ZONE 3": "qz3",
    "Al Nafel Park": "anp",
    "Al Khuzama Zone -2": "akz2",
    "Al Khuzama Zone -1": "akz1",
    "Road A1 - Al Khuzama": "raak",
    "Qetaifan North Park": "qnp",
    "Seef Lusail North": "sln",
}

# Fix for earlier typos: ensure anp placeholders use 'anp' not 'anp1' etc.
# For final placeholders we'll emit (1<code>t) and (2<code>t) where appropriate.

# -------------------------
# Build placeholder map from records
# -------------------------
def process_new_record_updates(day, new_json_files):
    """
    Parse newly downloaded JSON files (only those in new_json_files),
    extract record_update entries, and construct a text_map:
      - per-cage placeholders: e.g. "1c471": "3"
      - opposite shift placeholders: set to "0" if not present
      - per-place totals: e.g. "(1sp_total)", "(1spm)", "(1spl)"
    Save the mapping to OUTPUT_DIR/updates_<day>.json for doc_utils to consume.
    """
    if not new_json_files:
        return None

    data_dir = os.path.join(LOCAL_DIR, day, "data")
    photo_dir = os.path.join(LOCAL_DIR, day, "photos")

    # initialize mappings
    # hold per-place->shift->cage->value
    place_shift_cage = {}
    # hold per-place per-shift myna/local sums
    place_shift_myna = {}
    place_shift_local = {}

    # initialize for all known places and both shifts
    for place, cages in PLACE_CAGE_MAP.items():
        place_shift_cage.setdefault(place, {"1": {}, "2": {}})
        place_shift_myna.setdefault(place, {"1": 0, "2": 0})
        place_shift_local.setdefault(place, {"1": 0, "2": 0})
        # default zero for all cages (we'll set explicit values when records exist)
        for c in cages:
            place_shift_cage[place]["1"][str(c)] = None
            place_shift_cage[place]["2"][str(c)] = None

    # parse only the newly downloaded files (they are filenames like record_001.json)
    for fname in new_json_files:
        fpath = os.path.join(data_dir, fname)
        if not os.path.exists(fpath):
            continue
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                rec = json.load(f)
        except Exception as e:
            log(f"Failed to read JSON {fpath}: {e}")
            continue

        if rec.get("type") != "record_update":
            continue

        shift = str(rec.get("shift", "")).strip()
        cage = str(rec.get("cage_number", "")).strip()
        location = rec.get("location", "").strip()
        # numeric counts
        try:
            myna = int(rec.get("myna_captured", 0))
        except Exception:
            try:
                myna = int(str(rec.get("myna_captured", "0")).strip())
            except:
                myna = 0
        try:
            local_released = int(rec.get("local_released", 0))
        except Exception:
            try:
                local_released = int(str(rec.get("local_released", "0")).strip())
            except:
                local_released = 0

        total_birds = myna + local_released

        # Find canonical place from provided location string. Use direct match, else try substring match.
        place_key = None
        if location in PLACE_CAGE_MAP:
            place_key = location
        else:
            # try case-insensitive substring match
            for pk in PLACE_CAGE_MAP.keys():
                if pk.lower() in location.lower() or location.lower() in pk.lower():
                    place_key = pk
                    break

        if not place_key:
            log(f"Unknown place '{location}' in record {fname} — skipping.")
            continue

        # ensure cage belongs to place (if not, still accept but warn)
        if int(cage) not in PLACE_CAGE_MAP[place_key]:
            log(f"Cage {cage} for place {place_key} not in known cage list — accepting but check template mapping.")

        # store the value for that cage & shift
        place_shift_cage[place_key].setdefault(shift, {})[cage] = total_birds
        # accumulate myna/local for totals
        place_shift_myna[place_key][shift] += myna
        place_shift_local[place_key][shift] += local_released

        # set opposite shift cage to 0 if not set (explicit)
        other = "2" if shift == "1" else "1"
        if place_shift_cage[place_key][other].get(cage) is None:
            place_shift_cage[place_key][other][cage] = 0

    # Build text_map of placeholders -> values
    text_map = {}

    for place, cages in PLACE_CAGE_MAP.items():
        code = PLACE_CODE_MAP.get(place)
        if not code:
            # skip places without code mapping (shouldn't happen)
            log(f"No placeholder code for place '{place}' — skipping totals placeholders.")
            continue

        # per cage placeholders for both shifts
        for shift in ("1", "2"):
            for c in cages:
                key = f"{shift}c{c}"
                val = place_shift_cage.get(place, {}).get(shift, {}).get(str(c))
                # if still None => not updated by either shift -> keep empty string (so doc keeps previous or blank)
                if val is None:
                    # leave empty string so doc shows blank instead of 0
                    text_map[key] = ""
                else:
                    text_map[key] = str(val)

        # totals placeholders - follow your naming scheme: e.g. (2sp_total), (2spm), (2spl)
        # for shift 1 and 2
        for shift in ("1", "2"):
            t_prefix = f"({shift}{code}t)"      # total
            m_prefix = f"({shift}{code}m)"      # myna
            l_prefix = f"({shift}{code}l)"      # local

            total_val = 0
            # sum up per-cage numbers for this place & shift where not None
            for c in cages:
                v = place_shift_cage.get(place, {}).get(shift, {}).get(str(c))
                if v is None or v == "":
                    # treat None as 0 for totals if shift hasn't recorded anything yet
                    continue
                try:
                    total_val += int(v)
                except:
                    pass

            # fallback to the separate myna/local accumulators if available
            myna_val = place_shift_myna.get(place, {}).get(shift, 0)
            local_val = place_shift_local.get(place, {}).get(shift, 0)

            # Ensure myna+local <= total_val; prefer computed sum if non-zero, else fallback to myna/local
            if total_val == 0 and (myna_val or local_val):
                total_val = myna_val + local_val

            text_map[t_prefix] = str(total_val)
            text_map[m_prefix] = str(myna_val)
            text_map[l_prefix] = str(local_val)

    # persist text_map to OUTPUT_DIR for doc_utils to consume
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    updates_path = os.path.join(OUTPUT_DIR, f"updates_{day}.json")
    try:
        with open(updates_path, "w", encoding="utf-8") as uf:
            json.dump(text_map, uf, indent=2)
        log(f"Wrote updates mapping to {updates_path} (contains {len(text_map)} entries)")
    except Exception as e:
        log(f"Failed to write updates mapping: {e}")
        return None

    return updates_path

# -------------------------
# Main sync_day (modified to track newly-downloaded JSON files)
# -------------------------
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

    # track newly downloaded JSONs (filenames)
    new_json_files = []

    for f in server_data:
        if f not in known_data:
            if download_file(BASE_URL + f"{day}/data/{f}", os.path.join(data_dir, f)):
                download_db[day]["data"].append(f)
                save_download_db(download_db)
                new_data = True
                new_json_files.append(f)

    for f in server_photos:
        if f not in known_photos:
            if download_file(BASE_URL + f"{day}/photos/{f}", os.path.join(photos_dir, f)):
                download_db[day]["photos"].append(f)
                save_download_db(download_db)
                new_photos = True

    # Process new record_update entries (only for newly-downloaded JSON files)
    if new_json_files:
        updates_path = process_new_record_updates(day, new_json_files)
        if updates_path:
            log(f"Record updates processed; mapping at: {updates_path}")
        else:
            log("No record_update entries found in newly downloaded JSONs or processing failed.")

    # create report after sync (this will still create the partial report with shift sign images)
    if new_data or new_photos:
        partial_path = create_partial_report_with_shift_signs(day)
        if partial_path is None:
            log("Partial report creation failed.")
        else:
            if download_db.get(day, {}).get("finalized"):
                log(f"{day} already finalized — skipping finalization.")
            else:
                # Note: check_report_ready expects to examine local records. It will finalize only when ready.
                if check_report_ready(day, load_day_records_local):
                    final_path = finalize_report(day, partial_docx_path=partial_path)
                    if final_path:
                        log(f"Report finalized: {final_path}")
                    else:
                        log("Finalization attempt failed.")
        if len(server_photos) >= 10:
            log(f"Reached limit, deleting server files for {day}")
            delete_from_server(day)                
    else:
        log("No new files – report unchanged.")

    # if len(server_data) + len(server_photos) >= 10:
    #     log(f"Reached limit, deleting server files for {day}")
    #     delete_from_server(day)
