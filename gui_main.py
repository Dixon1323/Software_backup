import tkinter as tk
from tkinter import ttk, filedialog
import threading
import json
import os
import time
from datetime import datetime
import subprocess

from PIL import Image, ImageDraw
import pystray
from logger import log
from config import SETTINGS_FILE
from plyer import notification


# =============================
# GUI CLASS
# =============================
class SyncGUI(tk.Tk):
    def __init__(self):
        super().__init__()


        self.title("Word Report Sync System")
        self.geometry("650x430")
        self.resizable(False, False)

        # Load settings
        self.settings = self.load_settings()

        import config
        path = self.settings.get("REPORTS_DIR", "")
        config.OUTPUT_DIR = path 

        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        self.tray_icon = None
        self.tray_thread = None

        # Prepare state variables
        self.sync_running = False
        self.sync_thread = None
        self.paused = False
        self.last_shift1 = None
        self.last_shift2 = None

        # ------------------------
        # TITLE
        # ------------------------
        tk.Label(self, text="Word Report Sync System", font=("Arial", 18, "bold")).pack(pady=10)

        # ------------------------
        # CONTROL BUTTONS
        # ------------------------
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=8)

        self.start_stop_btn = tk.Button(btn_frame, text="Stop Sync", width=15, command=self.toggle_start_stop)
        self.start_stop_btn.grid(row=0, column=0, padx=10)

        self.pause_resume_btn = tk.Button(btn_frame, text="Pause", width=15, command=self.toggle_pause_resume)
        self.pause_resume_btn.grid(row=0, column=1, padx=10)

        # ------------------------
        # SHIFT STATUS
        # ------------------------
        self.shift1_status = tk.Label(self, text="Shift-1 Status: Not started", font=("Arial", 12), fg="blue")
        self.shift1_status.pack(pady=(15, 5))

        self.shift2_status = tk.Label(self, text="Shift-2 Status: Not started", font=("Arial", 12), fg="blue")
        self.shift2_status.pack(pady=(0, 15))

        # ------------------------
        # PROGRESS BAR
        # ------------------------
        self.progress_label = tk.Label(self, text="Progress: 0 / 178 Locations")
        self.progress_label.pack()

        self.progress = ttk.Progressbar(self, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=5)

        # ------------------------
        # SETTINGS
        # ------------------------
        settings_frame = tk.Frame(self)
        settings_frame.pack(pady=(10, 5))

        tk.Label(settings_frame, text="Loop Interval (sec):").grid(row=0, column=0, padx=5)
        self.interval_entry = tk.Entry(settings_frame, width=8)
        self.interval_entry.grid(row=0, column=1)
        self.interval_entry.insert(0, str(self.settings["LOOP_INTERVAL"]))

        tk.Button(settings_frame, text="Save Interval", command=self.save_interval).grid(row=0, column=2, padx=5)

        # Select Report Folder
        tk.Button(settings_frame, text="Select Reports Folder", command=self.select_reports_folder).grid(row=1, column=0, columnspan=3, pady=5)

        # Enable/Disable notifications
        self.enable_notifications = tk.BooleanVar(value=self.settings["ENABLE_NOTIFICATIONS"])
        tk.Checkbutton(self, text="Enable Notifications", variable=self.enable_notifications).pack()

        # ------------------------
        # FOOTER BUTTONS
        # ------------------------
        bottom_frame = tk.Frame(self)
        bottom_frame.pack(pady=15)

        tk.Button(bottom_frame, text="Open Reports Folder", width=20, command=self.open_reports_folder).grid(row=0, column=0, padx=5)
        tk.Button(bottom_frame, text="Open Last Report", width=20, command=self.open_last_report).grid(row=0, column=1, padx=5)
        #tk.Button(bottom_frame, text="Test Notification", width=18,command=lambda: self.notify("TEST", "If you see this, notifications work")).grid(row=0, column=2, padx=5)

        # ------------------------
        # AUTO-START SYNC
        # ------------------------
        self.after(200, self.start_sync)

        # Background UI updater
        self.after(1000, self.update_ui_loop)


    def create_tray_icon(self):
        # Create simple tray icon dynamically
        image = Image.new('RGB', (64, 64), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((0, 0, 64, 64), fill="blue")

        menu = pystray.Menu(
            pystray.MenuItem("Show", self.show_window),
            pystray.MenuItem("Exit", self.exit_app)
        )

        self.tray_icon = pystray.Icon("Daily Sync System", image, "Daily Sync System", menu)


    def minimize_to_tray(self):
        self.withdraw()  # Hide window

        if self.tray_icon is None:
            self.create_tray_icon()

            def run_tray():
                self.tray_icon.run()

            self.tray_thread = threading.Thread(target=run_tray, daemon=True)
            self.tray_thread.start()


    def show_window(self, icon=None, item=None):
        self.deiconify()     # Show window
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
    



    def exit_app(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
        self.destroy()
        os._exit(0)   # Force kill threads



    # ============================================================
    # LOAD SETTINGS
    # ============================================================
    def load_settings(self):
        if not os.path.exists(SETTINGS_FILE):
            return {
                "REPORTS_DIR": "",
                "LOOP_INTERVAL": 10,
                "ENABLE_NOTIFICATIONS": True
            }
        try:
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
        except:
            return {
                "REPORTS_DIR": "",
                "LOOP_INTERVAL": 10,
                "ENABLE_NOTIFICATIONS": True
            }


    def sync_loop(self):
        from config import BASE_URL
        from http_utils import safe_request
        from sync_day import sync_day

        while self.sync_running:
            import config

            if not config.OUTPUT_DIR or config.OUTPUT_DIR.strip() == "":
                log("Sync paused — report folder missing.")
                time.sleep(2)
                continue
            if not self.paused:
                # Fetch available date folders
                try:
                    res = safe_request(BASE_URL + "list_dates")
                    dates = res.json() if res else []
                except:
                    dates = []

                # Sync each date
                for day in dates:
                    if not self.sync_running:
                        break   # stopped while processing
                    if not self.paused:
                        sync_day(day)

            # Sleep based on config but GUI-driven
            time.sleep(self.settings["LOOP_INTERVAL"])

            
    # ============================================================
    # SAVE SETTINGS
    # ============================================================
    def save_settings(self):
        self.settings["ENABLE_NOTIFICATIONS"] = self.enable_notifications.get()
        with open(SETTINGS_FILE, "w") as f:
            json.dump(self.settings, f, indent=4)

    # ============================================================
    # LOOP INTERVAL
    # ============================================================
    def save_interval(self):
        try:
            self.settings["LOOP_INTERVAL"] = int(self.interval_entry.get())
            self.save_settings()
            log(f"Loop interval updated to {self.settings['LOOP_INTERVAL']} sec")
        except:
            log("Invalid interval entered!")

    # ============================================================
    # SELECT REPORT FOLDER
    # ============================================================
    def select_reports_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.settings["REPORTS_DIR"] = folder
            self.save_settings()

            import config
            config.OUTPUT_DIR = folder  # ← apply immediately

            log(f"Reports directory set to: {folder}")

    # ============================================================
    # NOTIFICATIONS (SAFE THREADED)
    # ============================================================
    def notify(self, title, msg):
        # do nothing if user disabled notifications
        try:
            if not getattr(self, "enable_notifications", tk.BooleanVar(value=True)).get():
                log(f"notify() skipped because notifications disabled: {title} - {msg}")
                return
        except Exception:
            # defensive fallback
            log("notify(): could not read enable_notifications var, continuing")

        def worker():
            log(f"notify(): attempt -> {title} | {msg}")
            # Primary: plyer
            try:
                notification.notify(
                    title=title,
                    message=msg,
                    app_name="Word Report Sync System",
                    timeout=5
                )
                log("notify(): plyer.notify succeeded")
                return
            except Exception as e:
                log(f"notify(): plyer failed: {e}")

            # Fallback 1: win10toast if available (non-blocking)
            try:
                from win10toast import ToastNotifier
                tn = ToastNotifier()
                tn.show_toast(title, msg, duration=5, threaded=True)
                log("notify(): win10toast succeeded")
                return
            except Exception as e:
                log(f"notify(): win10toast failed: {e}")

            # Fallback 2: simple tkinter messagebox (guaranteed visible)
            try:
                import tkinter.messagebox as mb
                # Must call messagebox on main thread: schedule via after()
                def show_mb():
                    try:
                        mb.showinfo(title, msg)
                        log("notify(): messagebox shown as fallback")
                    except Exception as e2:
                        log(f"notify(): messagebox failed: {e2}")
                # schedule on mainloop to avoid thread UI issues
                try:
                    self.after(0, show_mb)
                    log("notify(): scheduled messagebox via after(0,... )")
                except Exception as e3:
                    log(f"notify(): scheduling messagebox failed: {e3}")
                return
            except Exception as e:
                log(f"notify(): final fallback failed: {e}")
                return

        threading.Thread(target=worker, daemon=True).start()

    # ============================================================
    # START/STOP SYNC
    # ============================================================
    def start_sync(self):
        import config

        # BLOCK SYNC if user has NOT selected a folder
        if not config.OUTPUT_DIR or config.OUTPUT_DIR.strip() == "":
            log("Cannot start sync: No report folder selected!")
            self.notify("Sync Error", "Please select a Reports Folder first.")
            return
        
        if not self.sync_running:
            self.sync_running = True

            # Start GUI-controlled sync loop
            self.sync_thread = threading.Thread(target=self.sync_loop, daemon=True)
            self.sync_thread.start()

            self.start_stop_btn.config(text="Stop Sync")
            log("Sync started")


    def stop_sync(self):
        self.sync_running = False
        self.start_stop_btn.config(text="Start Sync")
        log("Sync stopped")

    
    


    def toggle_start_stop(self):
        if self.sync_running:
            self.stop_sync()
        else:
            self.start_sync()

    # ============================================================
    # PAUSE/RESUME
    # ============================================================
    def toggle_pause_resume(self):
        self.paused = not self.paused
        self.pause_resume_btn.config(text="Resume" if self.paused else "Pause")
        log("Sync paused" if self.paused else "Sync resumed")

    # ============================================================
    # UPDATE UI LOOP
    # ============================================================
    def update_ui_loop(self):
        self.update_progress()
        self.detect_shift_updates()
        self.after(1000, self.update_ui_loop)

    # ============================================================
    # PROGRESS (READ FROM downloaded_files.json)
    # ============================================================
    def update_progress(self):
        from config import DOWNLOADED_DB

        if not os.path.exists(DOWNLOADED_DB):
            return

        try:
            with open(DOWNLOADED_DB, "r") as f:
                db = json.load(f)
        except:
            return

        today = datetime.now().strftime("%Y-%m-%d")

        if today not in db:
            return

        data_count = len(db[today].get("data", []))

        # update label
        self.progress_label.config(
            text=f"Progress: {data_count} / 178 Locations"
        )

        # Set progress bar by percentage of photos (main target)
        percent = int((data_count / 178) * 100)
        self.progress["value"] = percent

    # ============================================================
    # SHIFT STATUS MONITOR
    # ============================================================
    def detect_shift_updates(self):
        """
        Improved shift detector:
        - Scans all downloaded JSON records for today
        - For each shift chooses the latest event (by timestamp, fallback to file mtime)
        - Updates UI once per shift if the state changed
        - Does NOT spam notifications on startup (first-run only sets state)
        """
        from config import DOWNLOADED_DB

        if not os.path.exists(DOWNLOADED_DB):
            return

        try:
            with open(DOWNLOADED_DB, "r") as f:
                db = json.load(f)
        except Exception:
            return

        today = datetime.now().strftime("%Y-%m-%d")
        if today not in db:
            return

        # Collect candidate events for each shift
        shift_events = {"1": [], "2": []}

        for rec_name in db[today].get("data", []):
            if not rec_name.endswith(".json"):
                continue

            full_path = os.path.join("sync", "records", today, "data", rec_name)
            if not os.path.exists(full_path):
                continue

            try:
                with open(full_path, "r", encoding="utf-8") as f:
                    rec = json.load(f)
            except Exception:
                # corrupt or unreadable file
                continue

            rtype = rec.get("type")
            shift = str(rec.get("shift", "")).strip()

            if rtype not in ("start_shift", "end_shift"):
                continue
            if shift not in ("1", "2"):
                continue

            # determine event time: try record timestamp, else file mtime
            try:
                mtime = os.path.getmtime(full_path)
                event_time = datetime.fromtimestamp(mtime)
            except:
                event_time = datetime.now()

            shift_events[shift].append({
                "type": rtype,
                "time": event_time,
                "file": rec_name
            })

        # For each shift pick the latest event (if any) and update UI once
        for shift in ("1", "2"):
            events = shift_events[shift]
            if not events:
                continue

            # sort by time ascending, take last
            events.sort(key=lambda e: e["time"])
            latest = events[-1]
            new_state = "IN" if latest["type"] == "start_shift" else "OUT"
            display_state = "Signed IN" if new_state == "IN" else "Signed OUT"
            color = "green" if new_state == "IN" else "red"

            # SHIFT 1
            if shift == "1":
                # If we never had a state (first run), set it silently (no toast)
                if self.last_shift1 is None and new_state == "IN":
                    # First state is IN → notify
                    self.shift1_status.config(text=f"Shift-1: {display_state}", fg=color)
                    self.last_shift1 = new_state
                    if self.enable_notifications.get():
                        self.notify("Shift Update", f"Shift-1 {display_state}")
                elif self.last_shift1 is None:
                    # If first state is OUT, set silently
                    self.shift1_status.config(text=f"Shift-1: {display_state}", fg=color)
                    self.last_shift1 = new_state
                elif self.last_shift1 != new_state:
                    # normal change
                    self.shift1_status.config(text=f"Shift-1: {display_state}", fg=color)
                    self.last_shift1 = new_state
                    if self.enable_notifications.get():
                        self.notify("Shift Update", f"Shift-1 {display_state}")
            # SHIFT 2
            else:
                if self.last_shift2 is None and new_state == "IN":
                    # First state is IN → notify
                    self.shift2_status.config(text=f"Shift-2: {display_state}", fg=color)
                    self.last_shift2 = new_state
                    if self.enable_notifications.get():
                        self.notify("Shift Update", f"Shift-2 {display_state}")
                elif self.last_shift2 is None:
                    # If first state is OUT, set silently
                    self.shift2_status.config(text=f"Shift-2: {display_state}", fg=color)
                    self.last_shift2 = new_state
                elif self.last_shift2 != new_state:
                    # normal change
                    self.shift2_status.config(text=f"Shift-2: {display_state}", fg=color)
                    self.last_shift2 = new_state
                    if self.enable_notifications.get():
                        self.notify("Shift Update", f"Shift-2 {display_state}")


    # ============================================================
    # FILE OPENING
    # ============================================================
    def open_reports_folder(self):
        import config

        root_folder = config.OUTPUT_DIR
        if not root_folder or not os.path.isdir(root_folder):
            log("Open Reports Folder failed: OUTPUT_DIR invalid.")
            return

        final_folder = os.path.join(root_folder, "final")

        if not os.path.isdir(final_folder):
            log("Open Reports Folder: 'final' folder not found.")
            return

        os.startfile(final_folder)
        log(f"Opened reports folder: {final_folder}")

    def open_last_report(self):
        import config

        root_folder = config.OUTPUT_DIR
        if not root_folder or not os.path.isdir(root_folder):
            log("Open Last Report failed: OUTPUT_DIR invalid.")
            return

        final_folder = os.path.join(root_folder, "final")

        if not os.path.isdir(final_folder):
            log("Open Last Report failed: 'final' folder missing.")
            return

        try:
            files = [
                f for f in os.listdir(final_folder)
                if f.lower().endswith(".docx")
            ]
        except:
            log("Could not list files in final/")
            return

        if not files:
            log("No reports found inside final/")
            return

        # Most recent file
        last_file = max(
            files,
            key=lambda f: os.path.getmtime(os.path.join(final_folder, f))
        )

        path = os.path.join(final_folder, last_file)
        os.startfile(path)
        log(f"Opened last report: {path}")


# ============================================================
# RUN GUI
# ============================================================
if __name__ == "__main__":
    app = SyncGUI()
    app.mainloop()
