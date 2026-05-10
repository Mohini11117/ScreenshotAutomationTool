import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from automation import run_automation
import time

# =========================================================
# COLORS
# =========================================================
BG_COLOR = "#081229"
CARD_COLOR = "#1A2740"
ACCENT_BLUE = "#1D4ED8"
ACCENT_BLUE_HOVER = "#1E40AF"
TEXT_COLOR = "#FFFFFF"
SECONDARY_TEXT = "#94A3B8"
LOG_BG = "#020617"
SUCCESS_COLOR = "#22C55E"
WARNING_COLOR = "#F59E0B"
ERROR_COLOR = "#EF4444"


class AutomationApp:

    def __init__(self, root):

        self.root = root

        # =====================================================
        # WINDOW SETTINGS
        # =====================================================
        self.root.title("ETM Automation")

        window_width = 1050
        window_height = 720

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)

        self.root.geometry(
            f"{window_width}x{window_height}+{x}+{y}"
        )

        self.root.configure(bg=BG_COLOR)
        self.root.resizable(False, False)

        # =====================================================
        # VARIABLES
        # =====================================================
        self.output_path = tk.StringVar()
        self.excel_path = tk.StringVar()

        self.start_time = None
        self.timer_running = False

        # =====================================================
        # STYLE
        # =====================================================
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor="#111827",
            background=ACCENT_BLUE,
            bordercolor="#111827",
            lightcolor=ACCENT_BLUE,
            darkcolor=ACCENT_BLUE,
            thickness=14
        )

        # =====================================================
        # HEADER
        # =====================================================
        header_frame = tk.Frame(
            root,
            bg=CARD_COLOR,
            height=70
        )

        header_frame.pack(fill="x")

        tk.Label(
            header_frame,
            text="ETM AUTOMATION",
            font=("Segoe UI", 22, "bold"),
            bg=CARD_COLOR,
            fg=TEXT_COLOR
        ).pack(pady=(10, 0))

        tk.Label(
            header_frame,
            text="Automated Screenshot & Documentation Platform",
            font=("Segoe UI", 10),
            bg=CARD_COLOR,
            fg=SECONDARY_TEXT
        ).pack()

        # =====================================================
        # MAIN CONTAINER
        # =====================================================
        container = tk.Frame(
            root,
            bg=BG_COLOR
        )

        container.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=10
        )

        # =====================================================
        # INPUT CARD
        # =====================================================
        input_card = tk.Frame(
            container,
            bg=CARD_COLOR
        )

        input_card.pack(fill="x")

        # =====================================================
        # OUTPUT FOLDER
        # =====================================================
        tk.Label(
            input_card,
            text="Output Folder",
            font=("Segoe UI", 11, "bold"),
            bg=CARD_COLOR,
            fg=TEXT_COLOR
        ).pack(anchor="w", padx=18, pady=(15, 5))

        output_frame = tk.Frame(
            input_card,
            bg=CARD_COLOR
        )

        output_frame.pack(fill="x", padx=18)

        self.output_entry = tk.Entry(
            output_frame,
            textvariable=self.output_path,
            font=("Segoe UI", 10),
            bg="#020617",
            fg="white",
            insertbackground="white",
            relief="flat",
            bd=0
        )

        self.output_entry.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=8,
            padx=(0, 10)
        )

        browse_btn1 = tk.Button(
            output_frame,
            text="Browse",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT_BLUE,
            fg="white",
            activebackground=ACCENT_BLUE_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=18,
            pady=7,
            command=self.select_folder
        )

        browse_btn1.pack(side="right")

        # =====================================================
        # INPUT EXCEL
        # =====================================================
        tk.Label(
            input_card,
            text="Input Excel File",
            font=("Segoe UI", 11, "bold"),
            bg=CARD_COLOR,
            fg=TEXT_COLOR
        ).pack(anchor="w", padx=18, pady=(15, 5))

        excel_frame = tk.Frame(
            input_card,
            bg=CARD_COLOR
        )

        excel_frame.pack(
            fill="x",
            padx=18,
            pady=(0, 15)
        )

        self.excel_entry = tk.Entry(
            excel_frame,
            textvariable=self.excel_path,
            font=("Segoe UI", 10),
            bg="#020617",
            fg="white",
            insertbackground="white",
            relief="flat",
            bd=0
        )

        self.excel_entry.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=8,
            padx=(0, 10)
        )

        browse_btn2 = tk.Button(
            excel_frame,
            text="Browse",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT_BLUE,
            fg="white",
            activebackground=ACCENT_BLUE_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=18,
            pady=7,
            command=self.select_excel
        )

        browse_btn2.pack(side="right")

        # =====================================================
        # STATS SECTION
        # =====================================================
        stats_frame = tk.Frame(
            container,
            bg=BG_COLOR
        )

        stats_frame.pack(fill="x", pady=6)

        self.total_card = self.create_stat_card(
            stats_frame,
            "TOTAL",
            "0"
        )

        self.completed_card = self.create_stat_card(
            stats_frame,
            "COMPLETED",
            "0"
        )

        self.failed_card = self.create_stat_card(
            stats_frame,
            "FAILED",
            "0"
        )

        self.current_card = self.create_stat_card(
            stats_frame,
            "CURRENT ETM",
            "-"
        )

        # =====================================================
        # START BUTTON
        # =====================================================
        self.start_btn = tk.Button(
            container,
            text="▶ START AUTOMATION",
            font=("Segoe UI", 12, "bold"),
            bg=ACCENT_BLUE,
            fg="white",
            activebackground=ACCENT_BLUE_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=22,
            pady=10,
            command=self.start_thread
        )

        self.start_btn.pack(pady=(3, 8))

        # =====================================================
        # PROGRESS CARD
        # =====================================================
        progress_card = tk.Frame(
            container,
            bg=CARD_COLOR
        )

        progress_card.pack(fill="x")

        tk.Label(
            progress_card,
            text="Automation Progress",
            font=("Segoe UI", 11, "bold"),
            bg=CARD_COLOR,
            fg=TEXT_COLOR
        ).pack(anchor="w", padx=18, pady=(12, 8))

        self.progress = ttk.Progressbar(
            progress_card,
            style="Custom.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate"
        )

        self.progress.pack(
            fill="x",
            padx=18
        )

        self.progress_percent = tk.Label(
            progress_card,
            text="0%",
            font=("Segoe UI", 10, "bold"),
            bg=CARD_COLOR,
            fg=ACCENT_BLUE
        )

        self.progress_percent.pack(
            anchor="e",
            padx=18,
            pady=5
        )

        # =====================================================
        # TIMER + STATUS
        # =====================================================
        info_frame = tk.Frame(
            progress_card,
            bg=CARD_COLOR
        )

        info_frame.pack(
            fill="x",
            padx=18,
            pady=(0, 12)
        )

        self.timer_label = tk.Label(
            info_frame,
            text="Elapsed Time: 00:00:00",
            font=("Segoe UI", 10, "bold"),
            bg=CARD_COLOR,
            fg=WARNING_COLOR
        )

        self.timer_label.pack(side="left")

        self.status_label = tk.Label(
            info_frame,
            text="Status: Idle",
            font=("Segoe UI", 10),
            bg=CARD_COLOR,
            fg=SECONDARY_TEXT
        )

        self.status_label.pack(side="right")

        # =====================================================
        # LOG CARD
        # =====================================================
        log_card = tk.Frame(
            container,
            bg=CARD_COLOR
        )

        log_card.pack(
            fill="both",
            expand=True,
            pady=8
        )

        tk.Label(
            log_card,
            text="LIVE LOGS",
            font=("Segoe UI", 11, "bold"),
            bg=CARD_COLOR,
            fg=TEXT_COLOR
        ).pack(anchor="w", padx=18, pady=(12, 8))

        log_frame = tk.Frame(
            log_card,
            bg=CARD_COLOR
        )

        log_frame.pack(
            fill="both",
            expand=True,
            padx=18,
            pady=(0, 15)
        )

        self.log_box = tk.Text(
            log_frame,
            bg=LOG_BG,
            fg=SUCCESS_COLOR,
            insertbackground="white",
            relief="flat",
            bd=0,
            font=("Consolas", 10),
            padx=12,
            pady=12,
            height=8
        )

        self.log_box.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar = tk.Scrollbar(log_frame)

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.log_box.config(
            yscrollcommand=scrollbar.set
        )

        scrollbar.config(
            command=self.log_box.yview
        )

        # =====================================================
        # FOOTER
        # =====================================================
        footer = tk.Frame(
            root,
            bg="#020617",
            height=28
        )

        footer.pack(fill="x")

        tk.Label(
            footer,
            text="Ready  |  ETM Automation Tool  |  Version 2.0",
            bg="#020617",
            fg=SECONDARY_TEXT,
            font=("Segoe UI", 8)
        ).pack(pady=6)

    # =====================================================
    # CREATE STAT CARD
    # =====================================================
    def create_stat_card(self, parent, title, value):

        card = tk.Frame(
            parent,
            bg=CARD_COLOR,
            width=180,
            height=65
        )

        card.pack(
            side="left",
            expand=True,
            fill="x",
            padx=4
        )

        card.pack_propagate(False)

        tk.Label(
            card,
            text=title,
            font=("Segoe UI", 9, "bold"),
            bg=CARD_COLOR,
            fg=SECONDARY_TEXT
        ).pack(pady=(8, 2))

        value_label = tk.Label(
            card,
            text=value,
            font=("Segoe UI", 16, "bold"),
            bg=CARD_COLOR,
            fg=TEXT_COLOR
        )

        value_label.pack()

        return value_label

    # =====================================================
    # SELECT FOLDER
    # =====================================================
    def select_folder(self):

        folder = filedialog.askdirectory()

        if folder:
            self.output_path.set(folder)

    # =====================================================
    # SELECT EXCEL
    # =====================================================
    def select_excel(self):

        file = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if file:
            self.excel_path.set(file)

    # =====================================================
    # UPDATE TIMER
    # =====================================================
    def update_timer(self):

        if self.timer_running:

            elapsed = int(time.time() - self.start_time)

            hours = elapsed // 3600
            minutes = (elapsed % 3600) // 60
            seconds = elapsed % 60

            self.timer_label.config(
                text=f"Elapsed Time: "
                     f"{hours:02}:{minutes:02}:{seconds:02}"
            )

            self.root.after(1000, self.update_timer)

    # =====================================================
    # LOG FUNCTION
    # =====================================================
    def log(self, message):

        self.status_label.config(text=message)

        self.log_box.insert(
            tk.END,
            message + "\n"
        )

        self.log_box.see(tk.END)

        if "Processing" in message:

            try:
                parts = message.split(":")[1].strip().split("/")

                current = int(parts[0])
                total = int(parts[1])

                self.progress["maximum"] = total
                self.progress["value"] = current

                percent = int((current / total) * 100)

                self.progress_percent.config(
                    text=f"{percent}%"
                )

                self.total_card.config(text=str(total))
                self.completed_card.config(text=str(current))

            except:
                pass

    # =====================================================
    # LOGIN POPUP
    # =====================================================
    def ask_login(self):

        messagebox.showinfo(
            "Login Required",
            "Please login in browser then click OK."
        )

    # =====================================================
    # START THREAD
    # =====================================================
    def start_thread(self):

        if not self.output_path.get():

            messagebox.showerror(
                "Error",
                "Please select output folder"
            )

            return

        if not self.excel_path.get():

            messagebox.showerror(
                "Error",
                "Please select Excel file"
            )

            return

        self.progress["value"] = 0

        self.start_btn.config(state="disabled")

        self.start_time = time.time()

        self.timer_running = True

        self.update_timer()

        thread = threading.Thread(
            target=self.run
        )

        thread.start()

    # =====================================================
    # RUN AUTOMATION
    # =====================================================
    def run(self):

        try:

            run_automation(
                self.output_path.get(),
                self.excel_path.get(),
                self.log,
                self.ask_login
            )

            elapsed = int(time.time() - self.start_time)

            hours = elapsed // 3600
            minutes = (elapsed % 3600) // 60
            seconds = elapsed % 60

            final_time = (
                f"{hours:02}:{minutes:02}:{seconds:02}"
            )

            messagebox.showinfo(
                "Success",
                f"Automation Completed!\n\n"
                f"Total Time: {final_time}"
            )

        except Exception as e:

            self.failed_card.config(text="1")

            messagebox.showerror(
                "Error",
                str(e)
            )

        finally:

            self.timer_running = False

            self.start_btn.config(state="normal")


# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":

    root = tk.Tk()

    app = AutomationApp(root)

    root.mainloop()