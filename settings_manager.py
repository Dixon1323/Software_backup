import os
import json

SETTINGS_FILE = "settings.json"

DEFAULT_SETTINGS = {
    "reports_folder": "",
    "loop_interval": 10,
    "notifications_enabled": True
}


def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        save_settings(DEFAULT_SETTINGS)
        return DEFAULT_SETTINGS.copy()

    try:
        with open(SETTINGS_FILE, "r") as f:
            data = json.load(f)

        # ensure all keys exist
        for k, v in DEFAULT_SETTINGS.items():
            if k not in data:
                data[k] = v

        return data

    except:
        save_settings(DEFAULT_SETTINGS)
        return DEFAULT_SETTINGS.copy()


def save_settings(data):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(data, f, indent=4)
