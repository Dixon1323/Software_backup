import os

APP_NAME = "DailySync"
APPDATA_DIR = os.path.join(os.getenv("APPDATA"), APP_NAME)

os.makedirs(APPDATA_DIR, exist_ok=True)


BASE_URL = "https://birdportal.pythonanywhere.com/records/"


SYNC_DIR = os.path.join(APPDATA_DIR, "sync")
LOCAL_DIR = os.path.join(SYNC_DIR, "records")
os.makedirs(LOCAL_DIR, exist_ok=True)

DOWNLOADED_DB = os.path.join(SYNC_DIR, "downloaded_files.json")
LOG_FILE = os.path.join(APPDATA_DIR, "sync.log")
SETTINGS_FILE = os.path.join(APPDATA_DIR, "settings.json")

DEFAULT_OUTPUT_DIR = os.path.join(LOCAL_DIR, "reports")
OUTPUT_DIR = DEFAULT_OUTPUT_DIR
os.makedirs(OUTPUT_DIR, exist_ok=True)

TEMPLATE_ORIG = os.path.join(os.path.dirname(__file__), "template.docx")

LOOP_INTERVAL = 10
MEDIA_EXT = ".png"
EMU_PER_PIXEL = 9525
