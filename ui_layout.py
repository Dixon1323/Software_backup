import customtkinter as ctk
from datetime import datetime


class ModernUI:
    """
    Modern / material-style UI wrapper.

    gui_main.py expects these attributes:
        - start_stop_btn
        - save_interval_btn
        - report_btn
        - open_reports_btn
        - open_last_btn
        - interval_entry
        - notifications_switch
        - progress
        - progress_label
        - shift1_status
        - shift2_status
        - status_pill
        - status_pill_bg
        - theme_switch   (new)
    """

    def __init__(self, root: ctk.CTk | ctk.CTkBaseClass):
        self.root = root

        # --- Global CTk look ---
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")  # you can try "green", "dark-blue", etc.

        self.root.title("Daily Sync System")
        self.root.geometry("900x520")
        self.root.minsize(820, 480)
        self.root.configure(bg="#020617")  # very dark background

        # =========================
        # OUTER "GLASS" CONTAINER
        # =========================
        outer = ctk.CTkFrame(
            self.root,
            fg_color="#020617",
            corner_radius=0
        )
        outer.pack(fill="both", expand=True, padx=12, pady=12)

        glass = ctk.CTkFrame(
            outer,
            corner_radius=26,
            fg_color="#020617"  # almost black, simulating glass panel
        )
        glass.pack(fill="both", expand=True, padx=2, pady=2)

        # =========================
        # TOP BAR
        # =========================
        top_bar = ctk.CTkFrame(glass, fg_color="transparent")
        top_bar.pack(fill="x", padx=18, pady=(14, 4))

        # App title with icon
        self.title_label = ctk.CTkLabel(
            top_bar,
            text="🕊️  Daily Sync System",
            font=("Segoe UI Semibold", 22),
            text_color="#e5e7eb"
        )
        self.title_label.pack(side="left")

        # Today label
        today_str = datetime.now().strftime("%d %b %Y")
        self.date_label = ctk.CTkLabel(
            top_bar,
            text=f"📅  {today_str}",
            font=("Segoe UI", 13),
            text_color="#9ca3af"
        )
        self.date_label.pack(side="left", padx=(14, 0))

        # Status pill background (rounded)
        self.status_pill_bg = ctk.CTkFrame(
            top_bar,
            width=120,
            height=30,
            fg_color="#1e293b",
            corner_radius=999
        )
        self.status_pill_bg.pack(side="right", padx=4)
        self.status_pill_bg.pack_propagate(False)

        # Actual pill text
        self.status_pill = ctk.CTkLabel(
            self.status_pill_bg,
            text="● Stopped",
            font=("Segoe UI", 13),
            text_color="#f97373"
        )
        self.status_pill.place(relx=0.5, rely=0.5, anchor="center")

        # internal pulse vars
        self._pulse_phase = 0
        self._start_status_pulse_loop()

        # =========================
        # MAIN 3-COLUMN AREA
        # =========================
        main = ctk.CTkFrame(glass, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=18, pady=(4, 14))

        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=2)
        main.grid_columnconfigure(2, weight=1)

        # ---------- LEFT: shift cards ----------
        left_panel = ctk.CTkFrame(main, corner_radius=20, fg_color="#020817")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=4)

        left_title = ctk.CTkLabel(
            left_panel,
            text="Shifts",
            font=("Segoe UI Semibold", 16),
            text_color="#e5e7eb"
        )
        left_title.pack(anchor="w", padx=16, pady=(14, 4))

        # Shift 1 card
        shift1_card = ctk.CTkFrame(left_panel, corner_radius=18, fg_color="#0b1120")
        shift1_card.pack(fill="x", padx=12, pady=(6, 4))

        shift1_header = ctk.CTkLabel(
            shift1_card,
            text="☀️  Shift 1",
            font=("Segoe UI Semibold", 14),
            text_color="#e5e7eb"
        )
        shift1_header.pack(anchor="w", padx=12, pady=(8, 2))

        self.shift1_status = ctk.CTkLabel(
            shift1_card,
            text="Status: Not started",
            font=("Segoe UI", 12),
            text_color="#60a5fa"
        )
        self.shift1_status.pack(anchor="w", padx=12, pady=(0, 10))

        # Shift 2 card
        shift2_card = ctk.CTkFrame(left_panel, corner_radius=18, fg_color="#0b1120")
        shift2_card.pack(fill="x", padx=12, pady=(8, 12))

        shift2_header = ctk.CTkLabel(
            shift2_card,
            text="🌙  Shift 2",
            font=("Segoe UI Semibold", 14),
            text_color="#e5e7eb"
        )
        shift2_header.pack(anchor="w", padx=12, pady=(8, 2))

        self.shift2_status = ctk.CTkLabel(
            shift2_card,
            text="Status: Not started",
            font=("Segoe UI", 12),
            text_color="#60a5fa"
        )
        self.shift2_status.pack(anchor="w", padx=12, pady=(0, 10))

        # ---------- CENTER: progress + controls ----------
        center_panel = ctk.CTkFrame(main, corner_radius=20, fg_color="#020817")
        center_panel.grid(row=0, column=1, sticky="nsew", padx=10, pady=4)

        center_panel.grid_rowconfigure(3, weight=1)  # allow some stretch

        center_title = ctk.CTkLabel(
            center_panel,
            text="Sync Overview",
            font=("Segoe UI Semibold", 16),
            text_color="#e5e7eb"
        )
        center_title.grid(row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(14, 10))

        # Progress label & bar
        self.progress_label = ctk.CTkLabel(
            center_panel,
            text="Progress: 0 / 178 Locations",
            font=("Segoe UI", 12),
            text_color="#9ca3af"
        )
        self.progress_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=18, pady=(0, 4))

        self.progress = ctk.CTkProgressBar(
            center_panel,
            orientation="horizontal",
            mode="determinate",
            height=14,
            corner_radius=999,
            fg_color="#020617",
            progress_color="#22c55e"
        )
        self.progress.set(0)
        self.progress.grid(row=2, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 12))

        # Start/Stop
        self.start_stop_btn = ctk.CTkButton(
            center_panel,
            text="▶️  Start Sync",
            font=("Segoe UI Semibold", 13),
            corner_radius=999,
            height=34
        )
        self.start_stop_btn.grid(row=3, column=0, sticky="w", padx=(18, 8), pady=(4, 4))

        # Divider line
        divider = ctk.CTkFrame(center_panel, height=1, fg_color="#111827")
        divider.grid(row=4, column=0, columnspan=2, sticky="ew", padx=18, pady=(12, 8))

        # Loop interval row
        loop_label = ctk.CTkLabel(
            center_panel,
            text="⏱️  Loop interval (sec):",
            font=("Segoe UI", 12),
            text_color="#d1d5db"
        )
        loop_label.grid(row=5, column=0, sticky="w", padx=(18, 4), pady=(6, 4))

        interval_frame = ctk.CTkFrame(center_panel, fg_color="transparent")
        interval_frame.grid(row=5, column=1, sticky="e", padx=(4, 18), pady=(6, 4))

        self.interval_entry = ctk.CTkEntry(
            interval_frame,
            width=70,
            justify="center",
            font=("Segoe UI", 12),
            corner_radius=999
        )
        self.interval_entry.pack(side="left", padx=(0, 6))

        self.save_interval_btn = ctk.CTkButton(
            interval_frame,
            text="💾 Save",
            width=70,
            height=30,
            font=("Segoe UI", 12),
            corner_radius=999,
            fg_color="#0ea5e9",
            hover_color="#0284c7"
        )
        self.save_interval_btn.pack(side="left")

        # Notifications toggle
        self.notifications_switch = ctk.CTkSwitch(
            center_panel,
            text="Desktop notifications",
            font=("Segoe UI", 12),
            text_color="#d1d5db",
            progress_color="#22c55e",
            button_color="#020617",
            button_hover_color="#020617"
        )
        self.notifications_switch.grid(row=6, column=0, columnspan=2, sticky="w", padx=18, pady=(10, 4))

        # Theme toggle
        self.theme_switch = ctk.CTkSwitch(
            center_panel,
            text="Dark mode",
            font=("Segoe UI", 12),
            text_color="#d1d5db",
            progress_color="#22c55e",
            button_color="#020617",
            button_hover_color="#020617",
            command=self._on_theme_toggle
        )
        self.theme_switch.grid(row=7, column=0, columnspan=2, sticky="w", padx=18, pady=(4, 12))
        self.theme_switch.select()  # default = dark mode

        # ---------- RIGHT: report controls ----------
        right_panel = ctk.CTkFrame(main, corner_radius=20, fg_color="#020817")
        right_panel.grid(row=0, column=2, sticky="nsew", padx=(10, 0), pady=4)

        right_title = ctk.CTkLabel(
            right_panel,
            text="Reports",
            font=("Segoe UI Semibold", 16),
            text_color="#e5e7eb"
        )
        right_title.pack(anchor="w", padx=16, pady=(14, 10))

        self.report_btn = ctk.CTkButton(
            right_panel,
            text="📂  Select Reports Folder",
            font=("Segoe UI", 12),
            corner_radius=999,
            height=34
        )
        self.report_btn.pack(fill="x", padx=16, pady=(4, 6))

        self.open_reports_btn = ctk.CTkButton(
            right_panel,
            text="📁  Open Reports Folder",
            font=("Segoe UI", 12),
            corner_radius=999,
            height=34,
            fg_color="#1e293b",
            hover_color="#111827",
        )
        self.open_reports_btn.pack(fill="x", padx=16, pady=6)

        self.open_last_btn = ctk.CTkButton(
            right_panel,
            text="📄  Open Last Report",
            font=("Segoe UI", 12),
            corner_radius=999,
            height=34,
            fg_color="#1e293b",
            hover_color="#111827",
        )
        self.open_last_btn.pack(fill="x", padx=16, pady=(6, 14))

        # Tiny hint at the bottom
        hint_label = ctk.CTkLabel(
            right_panel,
            text="ℹ️  Reports are saved inside a 'final' sub-folder.",
            font=("Segoe UI", 10),
            text_color="#6b7280",
            wraplength=220,
            justify="left"
        )
        hint_label.pack(anchor="w", padx=16, pady=(0, 10))

    # =========================
    # THEME TOGGLE HANDLER
    # =========================
    def _on_theme_toggle(self):
        if self.theme_switch.get():
            ctk.set_appearance_mode("dark")
            self.theme_switch.configure(text="Dark mode")
        else:
            ctk.set_appearance_mode("light")
            self.theme_switch.configure(text="Light mode")

    # =========================
    # STATUS PILL PULSE LOOP
    # =========================
    def _start_status_pulse_loop(self):
        self._animate_status_pill()

    def _animate_status_pill(self):
        """
        Simple pulse effect when status text contains 'Running'.
        Colors:
          - Running  → green pulse
          - Paused   → yellow static
          - Stopped  → red static
        """
        text = (self.status_pill.cget("text") or "").lower()

        if "running" in text:
            # pulsing green
            colors = ["#166534", "#16a34a", "#22c55e", "#16a34a"]
            idx = self._pulse_phase % len(colors)
            self.status_pill_bg.configure(fg_color=colors[idx])
            self.status_pill.configure(text_color="#f9fafb")
            self._pulse_phase += 1
        elif "paused" in text:
            # yellow
            self.status_pill_bg.configure(fg_color="#854d0e")
            self.status_pill.configure(text_color="#fde68a")
        else:
            # stopped or anything else → red
            self.status_pill_bg.configure(fg_color="#1e293b")
            self.status_pill.configure(text_color="#f97373")

        # schedule next frame
        self.root.after(450, self._animate_status_pill)
