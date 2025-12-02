# gui_worker.py
from PyQt6.QtCore import QThread, pyqtSignal
import time
import threading
import traceback

# We'll import the helpers used in your main loop
# main.get_available_dates and sync_day() from sync_day
try:
    from main import get_available_dates
except Exception:
    # fallback: in case main.py wasn't structured like that, try to get safe_request function
    get_available_dates = None

from sync_day import sync_day
import config


class SyncWorker(QThread):
    log_signal = pyqtSignal(str)
    status_signal = pyqtSignal(str)  # "running", "paused", "stopped", "error"

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pause_event = threading.Event()
        self._pause_event.clear()   # when set -> running; when clear -> paused
        self._stop_event = threading.Event()
        self._stop_event.clear()
        # default running at start
        self._pause_event.set()

    def run(self):
        self.log_signal.emit("=== GUI SYNC WORKER STARTED ===")
        self.status_signal.emit("running")
        loop_interval = getattr(config, "LOOP_INTERVAL", 10)
        # use get_available_dates from main if available, else call config.BASE_URL/list_dates manually
        local_get_dates = get_available_dates

        while not self._stop_event.is_set():
            try:
                if not self._pause_event.is_set():
                    # paused
                    self.status_signal.emit("paused")
                    self.log_signal.emit("Sync paused.")
                    # wait until resume or stop
                    while (not self._pause_event.is_set()) and (not self._stop_event.is_set()):
                        time.sleep(0.2)
                    if self._stop_event.is_set():
                        break
                    self.log_signal.emit("Sync resumed.")
                    self.status_signal.emit("running")

                # get dates list
                dates = []
                if local_get_dates:
                    try:
                        dates = local_get_dates()
                    except Exception as e:
                        self.log_signal.emit(f"get_available_dates() failed: {e}")
                        dates = []
                else:
                    # attempt to import and call safe_request endpoint fallback
                    try:
                        from http_utils import safe_request
                        import config as _cfg
                        resp = safe_request(_cfg.BASE_URL + "list_dates")
                        if resp is not None:
                            dates = resp.json()
                    except Exception as e:
                        self.log_signal.emit(f"Fallback get dates failed: {e}")

                if dates:
                    for day in dates:
                        if self._stop_event.is_set():
                            break
                        # check pause between days
                        while (not self._pause_event.is_set()) and (not self._stop_event.is_set()):
                            time.sleep(0.2)
                        try:
                            self.log_signal.emit(f"Syncing day: {day}")
                            sync_day(day)
                            self.log_signal.emit(f"Finished syncing day: {day}")
                        except Exception as e:
                            tb = traceback.format_exc()
                            self.log_signal.emit(f"Error while syncing {day}: {e}\n{tb}")
                else:
                    self.log_signal.emit("No dates available to sync.")

                # sleep loop interval but respect pause/stop quickly
                slept = 0.0
                while slept < loop_interval and not self._stop_event.is_set():
                    if not self._pause_event.is_set():
                        break
                    time.sleep(0.5)
                    slept += 0.5

            except Exception as e:
                tb = traceback.format_exc()
                self.log_signal.emit(f"Worker top-level exception: {e}\n{tb}")
                self.status_signal.emit("error")
                # short sleep to avoid tight crash loop
                time.sleep(2)

        self.status_signal.emit("stopped")
        self.log_signal.emit("Sync worker stopped.")

    def pause(self):
        self._pause_event.clear()
        self.log_signal.emit("Pause requested.")

    def resume(self):
        self._pause_event.set()
        self.log_signal.emit("Resume requested.")

    def stop(self):
        self._stop_event.set()
        # also unpause to let run exit quickly
        self._pause_event.set()
        self.log_signal.emit("Stop requested.")
