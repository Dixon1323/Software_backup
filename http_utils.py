import time
import requests
import os
from logger import log
from config import BASE_URL

def safe_request(url, retries=3):
    for i in range(retries):
        try:
            res = requests.get(url, timeout=10)
            if res.status_code == 200:
                return res
            else:
                log(f"Bad response {res.status_code}: {url}")
        except Exception as e:
            log(f"Network error accessing {url} ({i+1}/{retries}): {e}")
        time.sleep(2)
    return None

def download_file(url, local_path):
    res = safe_request(url)
    if res is None:
        log(f"FAILED downloading: {url}")
        return False
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    try:
        with open(local_path, "wb") as f:
            for chunk in res.iter_content(1024):
                f.write(chunk)
        log(f"Downloaded: {local_path}")
        return True
    except Exception as e:
        log(f"Failed saving download {local_path}: {e}")
        return False

def delete_from_server(day):
    url = BASE_URL + f"{day}/delete"
    try:
        res = requests.post(url)
        if res.status_code == 200:
            log(f"Server files deleted for {day}")
        else:
            log(f"Failed to delete server files for {day}: HTTP {res.status_code}")
    except Exception as e:
        log(f"Error deleting files on server for {day}: {e}")
