import os

# ----------------------------------------------
#   ROOT DIRECTORY (auto-detect project folder)
# ----------------------------------------------
ROOT = os.path.dirname(os.path.abspath(__file__))


# ----------------------------------------------
#   SERVER ENDPOINT
# ----------------------------------------------
BASE_URL = "https://birdportal.pythonanywhere.com/records/"


# ----------------------------------------------
#   LOCAL PATHS (AUTOMATICALLY CROSS-PLATFORM)
# ----------------------------------------------
SYNC_DIR = os.path.join(ROOT, "sync")
LOCAL_DIR = os.path.join(SYNC_DIR, "records")
OUTPUT_DIR = os.path.join(SYNC_DIR, "reports")
DOWNLOADED_DB = os.path.join(SYNC_DIR, "downloaded_files.json")
LOG_FILE = os.path.join(SYNC_DIR, "sync.log")
SETTINGS_FILE = os.path.join(ROOT, "settings.json")

# Template file inside project folder
TEMPLATE_ORIG = os.path.join(ROOT, "template.docx")


# ----------------------------------------------
#   RUNTIME SETTINGS
# ----------------------------------------------
LOOP_INTERVAL = 10  # seconds


# ----------------------------------------------
#   IMAGE / DOCX SETTINGS
# ----------------------------------------------
MEDIA_EXT = ".png"
EMU_PER_PIXEL = 9525
