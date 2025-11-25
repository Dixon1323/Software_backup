import os
import json
from config import DOWNLOADED_DB
from logger import log

def load_download_db():
    if not os.path.exists(DOWNLOADED_DB):
        return {}
    try:
        return json.load(open(DOWNLOADED_DB, encoding="utf-8"))
    except Exception as e:
        log(f"Failed loading download DB: {e}")
        return {}

def save_download_db(db):
    try:
        with open(DOWNLOADED_DB, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=4)
    except Exception as e:
        log(f"Failed saving DB: {e}")
