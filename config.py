import os


ROOT = os.path.dirname(os.path.abspath(__file__))
BASE_URL = "https://birdportal.pythonanywhere.com/records/"
SYNC_DIR = os.path.join(ROOT, "sync")
LOCAL_DIR = os.path.join(SYNC_DIR, "records")
OUTPUT_DIR = None
DOWNLOADED_DB = os.path.join(SYNC_DIR, "downloaded_files.json")
LOG_FILE = os.path.join(SYNC_DIR, "sync.log")
SETTINGS_FILE = os.path.join(ROOT, "settings.json")
TEMPLATE_ORIG = os.path.join(ROOT, "template.docx")
LOOP_INTERVAL = 10  
MEDIA_EXT = ".png"
EMU_PER_PIXEL = 9525
