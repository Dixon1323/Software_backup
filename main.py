import time
from logger import log
from config import LOOP_INTERVAL
from http_utils import safe_request
from sync_day import sync_day

def get_available_dates():
    res = safe_request(__import__('config').BASE_URL + "list_dates")
    if res is None:
        log("Could not get date folder list")
        return []
    try:
        return res.json()
    except Exception:
        return []

def main_loop():
    log("=== SYNC SYSTEM STARTED ===")
    while True:
        #log("Checking for new files...")
        dates = get_available_dates()
        if dates:
            for day in dates:
                sync_day(day)
        #log(f"Cycle complete. Sleeping {LOOP_INTERVAL} sec...")
        time.sleep(LOOP_INTERVAL)

if __name__ == "__main__":
    main_loop()