import time
from logger import log
from http_utils import safe_request
from sync_day import sync_day
from config import BASE_URL


def get_available_dates():
    """Fetch available date folders from server."""
    res = safe_request(BASE_URL + "list_dates")
    if res is None:
        log("Could not get date folder list")
        return []
    try:
        return res.json()
    except Exception:
        return []


def main_loop():
    """
    A SINGLE cycle of sync.
    GUI calls this repeatedly inside its own loop.
    """
    dates = get_available_dates()

    if dates:
        for day in dates:
            sync_day(day)
    else:
        log("No new dates available.")

    # DO NOT sleep here — GUI controls timing
