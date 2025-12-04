import tkinter as tk
from tkinter import filedialog
import threading
import json
import os
import time
from datetime import datetime
import subprocess

from logger import log
from config import SETTINGS_FILE
from plyer import notification
from PIL import Image
import pystray
from ui_layout import ModernUI


class SyncGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        try:
            self.iconbitmap(icon_path)
        except:
            pass

        # ------------------------------------
        # LOAD SETTINGS
        # ------------------------------------
        self.settings = self.load_settings()

        # Apply saved report folder to config
        import config
        saved_dir = self.settings.get("REPORTS_DIR", "")
        if saved_dir and os.path.exists(saved_dir):
            config.OUTPUT_DIR = saved_dir
        else:
            log("DEBUG: No saved REPORTS_DIR — using default AppData reports folder")

        self.tray_icon = None
        self.is_hidden_to_tray = False

        # intercept close button and minimize action
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        self.bind("<Unmap>", lambda e: self.minimize_to_tray() if self.state() == "iconic" else None)


        # ------------------------------------
        # STATE VARIABLES
        # ------------------------------------
        self.sync_running = False
        self.sync_thread = None
        self.paused = False
        self.last_shift1 = None
        self.last_shift2 = None

        # ------------------------------------
        # BUILD MODERN UI
        # ------------------------------------
        self.ui = ModernUI(self)

        import config
        if config.OUTPUT_DIR == config.DEFAULT_OUTPUT_DIR:
            self.update_status_pill("Needs Setup")
            self.notify("Setup Required", "Reporting temporarily using a default storage location.")


        # Attach events to UI buttons
        self.ui.start_stop_btn.configure(command=self.toggle_start_stop)
        # self.ui.pause_resume_btn.configure(command=self.toggle_pause_resume)
        self.ui.save_interval_btn.configure(command=self.save_interval)
        self.ui.report_btn.configure(command=self.select_reports_folder)
        self.ui.open_reports_btn.configure(command=self.open_reports_folder)
        self.ui.open_last_btn.configure(command=self.open_last_report)

        # Apply existing settings
        self.ui.interval_entry.insert(0, str(self.settings["LOOP_INTERVAL"]))
        self.ui.notifications_switch.select() if self.settings["ENABLE_NOTIFICATIONS"] else self.ui.notifications_switch.deselect()

        # ------------------------------------
        # AUTO-START SYNC
        # ------------------------------------
        self.after(200, self.start_sync)

        # Background UI update loop
        self.after(1000, self.update_ui_loop)

        # ------------------------------------
        # TRAY SUPPORT
        # ------------------------------------
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)


    # ============================================================
    def load_settings(self):
        if not os.path.exists(SETTINGS_FILE):
            return {"REPORTS_DIR": "", "LOOP_INTERVAL": 10, "ENABLE_NOTIFICATIONS": True}
        try:
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
        except:
            return {"REPORTS_DIR": "", "LOOP_INTERVAL": 10, "ENABLE_NOTIFICATIONS": True}


    def animate_pill(self):
        if getattr(self, "is_running_anim", False):
            current = self.ui.status_pill.cget("text_color")
            # Pulse between two greens
            pulse_green = "#ffffff"
            pulse_light = "#000000"
            self.ui.status_pill.configure(
                text_color=pulse_light if current == pulse_green else pulse_green
            )
            self.after(1000, self.animate_pill)

    def update_status_pill(self, state):
            if state == "Running":
                self.ui.status_pill.configure(text="● Running",text_color="#3adb76")  # green)
                self.ui.status_pill_bg.configure(fg_color="#0f5132")  # deep green
                self.is_running_anim = True
                self.animate_pill()
            elif state == "Paused":
                self.ui.status_pill.configure(text="● Paused",text_color="#f4c542")  # yellow)
                self.ui.status_pill_bg.configure(fg_color="#5a4b13")
                self.is_running_anim = False
            
            elif state == "Needs Setup":
                self.ui.status_pill.configure(text="● Setup Required", text_color="#fbbf24")
                self.ui.status_pill_bg.configure(fg_color="#78350f")
                self.is_running_anim = False

            else:  # Stopped
                self.ui.status_pill.configure(text="● Stopped",text_color="#f97373")  # red
                self.ui.status_pill_bg.configure(fg_color="#5b1d1d")
                self.is_running_anim = False

    # ============================================================
    # SETTINGS SAVE
    # ============================================================
    def save_settings(self):
        self.settings["ENABLE_NOTIFICATIONS"] = bool(self.ui.notifications_switch.get())
        with open(SETTINGS_FILE, "w") as f:
            json.dump(self.settings, f, indent=4)


    # ============================================================
    # START / STOP SYNC
    # ============================================================
    def start_sync(self):
        import config

        # Check if OUTPUT_DIR still default → warn and stop
        if config.OUTPUT_DIR == config.DEFAULT_OUTPUT_DIR:
            self.update_status_pill("Needs Setup")
            self.notify("Setup Required", "Please select a Reports Folder before syncing.")
            tk.messagebox.showwarning("Reports Folder Not Set",
                "Reports folder is currently using a default internal location.\n\n"
                "Please select a proper destination under: 'Select Reports Folder'")
            return
        if not self.sync_running:
            self.sync_running = True
            self.ui.start_stop_btn.configure(text="Stop Sync")

            # Start sync loop in background
            self.sync_thread = threading.Thread(target=self.sync_loop, daemon=True)
            self.sync_thread.start()

            log("Sync started")
            self.update_status_pill("Running")


    def stop_sync(self):
        self.sync_running = False
        self.ui.start_stop_btn.configure(text="Start Sync")
        log("Sync stopped")
        self.update_status_pill("Stopped")


    def toggle_start_stop(self):
        if self.sync_running:
            self.stop_sync()
            self.update_status_pill("Stopped")
        else:
            self.start_sync()
            self.update_status_pill("Running")


    # ============================================================
    # PAUSE / RESUME
    # ============================================================
    def toggle_pause_resume(self):
        self.paused = not self.paused
        self.ui.pause_resume_btn.configure(text="Resume" and self.update_status_pill("Running") if self.paused else "Pause" and self.update_status_pill("Paused"))
        log("Sync paused" if self.paused else "Sync resumed")


    # ============================================================
    # SYNC LOOP
    # ============================================================
    def sync_loop(self):
        from config import BASE_URL
        from http_utils import safe_request
        from sync_day import sync_day

        while self.sync_running:
            if not self.paused:
                try:
                    res = safe_request(BASE_URL + "list_dates")
                    dates = res.json() if res else []
                except:
                    dates = []

                for day in dates:
                    if not self.sync_running:
                        break
                    if not self.paused:
                        sync_day(day)

            time.sleep(self.settings["LOOP_INTERVAL"])


    # ============================================================
    # LOOP INTERVAL
    # ============================================================
    def save_interval(self):
        try:
            self.settings["LOOP_INTERVAL"] = int(self.ui.interval_entry.get())
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
            config.OUTPUT_DIR = folder

            log(f"Reports directory set to: {folder}")

            self.update_status_pill("Running")
            self.start_sync()

    # ============================================================
    # NOTIFICATIONS
    # ============================================================

    def notify(self, title, msg):
        if not bool(self.ui.notifications_switch.get()):
            return

        def worker():
            try:
                # Try WinNotify first
                from winotify import Notification, audio
                toast = Notification(
                    app_id="DPC Word Automation",
                    title=title,
                    msg=msg,
                    duration="short"
                )
                toast.set_audio(audio.Default, loop=False)
                toast.show()
                log(f"Notification: WinNotify SUCCESS")
                return
            except Exception as e:
                log(f"WinNotify failed: {e}")

            # Fallback to plyer
            try:
                from plyer import notification
                notification.notify(
                    title=title,
                    message=msg,
                    app_name="DPC Word Automation",
                    timeout=5
                )
                log(f"Notification: Plyer SUCCESS")
            except Exception as e:
                log(f"Notification error: {e}")

        threading.Thread(target=worker, daemon=True).start()

    # def notify(self, title, msg):
    #     if not bool(self.ui.notifications_switch.get()):
    #         return

    #     def worker():
    #         try:
    #             notification.notify(
    #                 title=title,
    #                 message=msg,
    #                 app_name="Daily Sync System",
    #                 timeout=5
    #             )
    #         except Exception as e:
    #             log(f"Notification error: {e}")

    #     threading.Thread(target=worker, daemon=True).start()


    # ============================================================
    # UI LOOP (progress + shift updates)
    # ============================================================
    def update_ui_loop(self):
        self.update_progress()
        self.detect_shift_updates()
        self.after(1000, self.update_ui_loop)


    # ============================================================
    # PROGRESS
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

        count = len(db[today].get("data", []))
        percent = int((count / 178) * 100)

        self.ui.progress_label.configure(text=f"Progress: {count} / 178 Locations")
        self.ui.progress.set(percent / 100)


    # ============================================================
    # SHIFT DETECTION (same logic as before)
    # ============================================================
    def detect_shift_updates(self):
            from config import DOWNLOADED_DB, LOCAL_DIR

            # log("DEBUG: detect_shift_updates() called")

            if not os.path.exists(DOWNLOADED_DB):
                # log("DEBUG: DOWNLOADED_DB does not exist → No shifts yet")
                return

            try:
                with open(DOWNLOADED_DB, "r") as f:
                    db = json.load(f)
            except Exception as e:
                # log(f"DEBUG: Failed to load download DB → {e}")
                return

            today = datetime.now().strftime("%Y-%m-%d")
            # log(f"DEBUG: Today = {today}")

            if today not in db:
                # log("DEBUG: No records found for today in DB")
                return

            shift_events = {"1": [], "2": []}
            day_folder = os.path.join(LOCAL_DIR, today, "data")
            # log(f"DEBUG: Local day folder → {day_folder}")

            if not os.path.exists(day_folder):
                # log("DEBUG: Local day folder does not exist → No shifts downloaded yet")
                return

            data_files = db[today].get("data", [])
            # log(f"DEBUG: Files listed in DB → {len(data_files)}")

            for name in data_files:
                if not name.endswith(".json"):
                    continue

                path = os.path.join(day_folder, name)
                # log(f"DEBUG: Checking file → {path}")

                if not os.path.exists(path):
                    # log("DEBUG: FILE NOT FOUND LOCALLY — CRITICAL PATH ISSUE")
                    continue

                try:
                    with open(path, "r") as f:
                        rec = json.load(f)
                except Exception as e:
                    # log(f"DEBUG: Failed to load JSON {name} → {e}")
                    continue

                rtype = rec.get("type")
                shift = str(rec.get("shift", ""))
                ts_str = rec.get("timestamp")

                # log(f"DEBUG: JSON {name} -> type={rtype}, shift={shift}, ts={ts_str}")

                if rtype not in ("start_shift", "end_shift"):
                    continue
                if shift not in ("1", "2"):
                    continue

                try:
                    event_time = datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
                except:
                    event_time = datetime.now()

                shift_events[shift].append({"type": rtype, "time": event_time})

            # SHOW FOUND SHIFT EVENTS
            # log(f"DEBUG: Shift event counts → Shift1={len(shift_events['1'])}, Shift2={len(shift_events['2'])}")

            for shift in ("1", "2"):
                if not shift_events[shift]:
                    # log(f"DEBUG: No events for shift {shift}")
                    continue

                latest = sorted(shift_events[shift], key=lambda e: e["time"])[-1]
                new_state = "IN" if latest["type"] == "start_shift" else "OUT"

                # log(f"DEBUG: Shift-{shift} latest state → {new_state}")

                label = self.ui.shift1_status if shift == "1" else self.ui.shift2_status
                last_state_attr = "last_shift1" if shift == "1" else "last_shift2"
                last = getattr(self, last_state_attr)

                if last is None:
                    # First shift event — update UI and notify
                    setattr(self, last_state_attr, new_state)
                    label.configure(
                        text=f"Signed {'IN' if new_state=='IN' else 'OUT'}",
                        text_color="green" if new_state == "IN" else "red"
                    )
                    self.notify(f"Shift Update", f"Shift-{shift} Signed {'IN' if new_state=='IN' else 'OUT'}")
                    # log(f"DEBUG: First notification -> Shift-{shift} {new_state}")
                    continue




    # ============================================================
    # OPEN FOLDERS
    # ============================================================
    def open_reports_folder(self):
        folder = self.settings["REPORTS_DIR"]
        if folder and os.path.exists(folder):
            os.startfile(folder)

    def open_last_report(self):
        folder = os.path.join(self.settings["REPORTS_DIR"], "final")
        if not os.path.exists(folder):
            return

        files = sorted(
            [f for f in os.listdir(folder) if f.endswith(".docx")],
            key=lambda x: os.path.getmtime(os.path.join(folder, x)),
            reverse=True
        )
        if files:
            os.startfile(os.path.join(folder, files[0]))


    # ============================================================
    # MINIMIZE TO TRAY
    # ============================================================
    def minimize_to_tray(self):
        if self.is_hidden_to_tray:
            return

        self.overrideredirect(False)
        self.withdraw()  # Hide window
        self.is_hidden_to_tray = True

        # Create tray icon image
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "app_icon.ico")
            image = Image.open(icon_path)
        except:
            # fallback: create a small blank icon
            image = Image.new('RGB', (64, 64), color='black')

        def restore(icon, item):
            self.restore_from_tray()

        def exit_app(icon, item):
            self.force_close()

        menu = pystray.Menu(
            pystray.MenuItem("Restore", restore),
            pystray.MenuItem("Exit", exit_app)
        )

        self.tray_icon = pystray.Icon(
            "Daily Sync System",
            image,
            "Daily Sync System",
            menu
        )

        threading.Thread(target=self.tray_icon.run, daemon=True).start()

        # Optional notification
        self.notify("Daily Sync System", "Running in background. Tray icon active.")


    def restore_from_tray(self):
        if self.tray_icon:
            self.tray_icon.stop()
        
        self.overrideredirect(True)
        self.is_hidden_to_tray = False
        self.tray_icon = None

        self.deiconify()  # Show window again
        self.after(10, self.lift)


    def force_close(self):
        self.sync_running = False

        if self.tray_icon:
            self.tray_icon.stop()

        self.destroy()
        os._exit(0)



if __name__ == "__main__":
    app = SyncGUI()
    app.mainloop()
