from __future__ import annotations

import tkinter as tk
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, ttk

from seed_builder import build_seed, workbook_question_count, write_runner


def rounded_rectangle(canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs):
    points = [
        x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
        x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
        x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)


class SeedBuilderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Seed Runner Builder")
        self.geometry("820x620")
        self.minsize(720, 560)
        self.configure(bg="#181818", padx=12, pady=12)

        self.folder_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / "SEED_RUNNER.py"))
        self._configure_style()
        self._build_ui()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background="#1e1e1e", foreground="#cccccc", font=("Segoe UI", 10))
        style.configure("Title.TLabel", background="#1e1e1e", foreground="#ffffff", font=("Segoe UI", 19, "bold"))
        style.configure("Subtitle.TLabel", background="#1e1e1e", foreground="#858585", font=("Segoe UI", 10))
        style.configure("Card.TLabelframe", background="#252526", foreground="#cccccc", bordercolor="#454545", relief="flat", borderwidth=1)
        style.configure("Card.TLabelframe.Label", background="#252526", foreground="#cccccc", font=("Segoe UI", 10, "bold"))
        style.configure("VSCode.TEntry", fieldbackground="#3c3c3c", foreground="#ffffff", insertcolor="#ffffff", borderwidth=1)
        style.configure("VSCode.TButton", background="#3c3c3c", foreground="#ffffff", borderwidth=0, padding=(14, 8))
        style.map("VSCode.TButton", background=[("active", "#505050")])
        style.configure("Import.TButton", background="#007acc", foreground="#ffffff", borderwidth=0, padding=(22, 10), font=("Segoe UI", 11, "bold"))
        style.map("Import.TButton", background=[("active", "#1488d4"), ("disabled", "#094771")])
        style.configure("Progress.Horizontal.TProgressbar", background="#007acc", troughcolor="#3c3c3c", bordercolor="#3c3c3c", lightcolor="#007acc", darkcolor="#007acc")

    def _build_ui(self) -> None:
        shell = tk.Canvas(self, bg="#181818", highlightthickness=0)
        shell.pack(fill="both", expand=True)
        main = tk.Frame(shell, bg="#252526", padx=24, pady=24, highlightbackground="#454545", highlightthickness=1)
        window = shell.create_window((12, 12), window=main, anchor="nw")

        def resize_shell(event):
            shell.coords(window, 12, 12)
            shell.itemconfigure(window, width=max(1, event.width - 24), height=max(1, event.height - 24))
            shell.delete("shell_border")
            rounded_rectangle(shell, 12, 12, event.width - 12, event.height - 12, 14, fill="#252526", outline="#454545", tags=("shell_border",))
            shell.tag_lower("shell_border", window)
        shell.bind("<Configure>", resize_shell)

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 18))
        ttk.Label(header, text="SEED RUNNER", style="Title.TLabel").pack(side="right")
        ttk.Label(header, text="יצירת קובץ הרצה מנתוני תרגול מתוך Excel", style="Subtitle.TLabel").pack(side="right", anchor="e", fill="x", expand=True, padx=(12, 0))

        self._add_path_row(main, "תיקיית תרגול", self.folder_var, self.choose_folder)
        self._add_path_row(main, "קובץ יעד", self.output_var, self.choose_output)

        log_canvas = tk.Canvas(main, bg="#252526", highlightthickness=0, height=190)
        log_canvas.pack(fill="both", expand=True, pady=(4, 12))
        log_canvas.create_text(14, 12, text="לוג ייבוא", fill="#cccccc", font=("Segoe UI", 10, "bold"), anchor="ne", tags=("log_border",))
        self.status = tk.Text(
            log_canvas, height=12, state="disabled", wrap="word", font=("Consolas", 10),
            bg="#1e1e1e", fg="#cccccc", insertbackground="#cccccc", relief="flat", padx=10, pady=10,
        )
        log_canvas.create_window((10, 36), window=self.status, anchor="nw", width=680, height=140, tags=("status_window",))
        def resize_log(event):
            log_canvas.delete("log_border")
            rounded_rectangle(log_canvas, 1, 1, event.width - 1, event.height - 1, 12, fill="#252526", outline="#454545", tags=("log_border",))
            log_canvas.coords("status_window", 10, 36)
            log_canvas.itemconfigure("status_window", width=max(1, event.width - 20), height=max(1, event.height - 46))
        log_canvas.bind("<Configure>", resize_log)

        self.status.tag_configure("info", foreground="#cccccc")
        self.status.tag_configure("success", foreground="#89d185")
        self.status.tag_configure("error", foreground="#f48771")
        self.status.tag_configure("warning", foreground="#dcdcaa")

        bottom = ttk.Frame(main)
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determinate", maximum=100, style="Progress.Horizontal.TProgressbar")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 18))
        self.progress["value"] = 0
        self.import_button = tk.Canvas(bottom, width=150, height=42, bg="#252526", highlightthickness=0)
        rounded_rectangle(self.import_button, 1, 1, 149, 41, 10, fill="#007acc", outline="#007acc", tags=("button_bg",))
        self.import_button.create_text(75, 21, text="יבא נתונים", fill="#ffffff", font=("Segoe UI", 11, "bold"), tags=("button_text",))
        self.import_button.bind("<Button-1>", lambda event: self.build())
        self.import_button.bind("<Enter>", lambda event: self.import_button.itemconfigure("button_bg", fill="#1488d4"))
        self.import_button.bind("<Leave>", lambda event: self.import_button.itemconfigure("button_bg", fill="#007acc"))
        self.import_button.pack(side="right", ipadx=18, ipady=6)
        self.log("מוכן לייבוא. בחרו תיקיית תרגול וקובץ יעד.", "info")

    def set_import_enabled(self, enabled: bool) -> None:
        color = "#007acc" if enabled else "#094771"
        self.import_button.itemconfigure("button_bg", fill=color, outline=color)
        self.import_button.unbind("<Button-1>")
        if enabled:
            self.import_button.bind("<Button-1>", lambda event: self.build())

    def _add_path_row(self, parent: ttk.Frame, label: str, variable: tk.StringVar, callback) -> None:
        frame = ttk.LabelFrame(parent, text=label, style="Card.TLabelframe", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        ttk.Entry(frame, textvariable=variable, style="VSCode.TEntry", justify="right").pack(side="right", fill="x", expand=True, padx=(10, 0))
        ttk.Button(frame, text="בחירה…", style="VSCode.TButton", command=callback).pack(side="right", padx=(10, 0))

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="בחר את תיקיית תרגול")
        if folder:
            self.folder_var.set(folder)
            self.log(f"נבחרה תיקיית תרגול: {folder}", "success")

    def choose_output(self) -> None:
        output = filedialog.asksaveasfilename(
            title="שמירת SEED_RUNNER.py",
            defaultextension=".py",
            initialfile="SEED_RUNNER.py",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
        )
        if output:
            self.output_var.set(output)
            self.log(f"נבחר יעד שמירה: {output}", "success")

    def log(self, text: str, level: str = "info") -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {"success": "[הצליח]", "error": "[נכשל]", "warning": "[אזהרה]"}.get(level, "[מידע]")
        self.status.configure(state="normal")
        self.status.insert("end", f"{timestamp} {prefix} {text}\n", (level,))
        self.status.configure(state="disabled")
        self.status.see("end")
        self.update_idletasks()

    def build(self) -> None:
        folder_text = self.folder_var.get().strip()
        output_text = self.output_var.get().strip()
        if not folder_text:
            self.log("יש לבחור תיקיית תרגול.", "error")
            return
        if not output_text:
            self.log("יש לבחור קובץ יעד.", "error")
            return

        try:
            self.set_import_enabled(False)
            self.progress["value"] = 5
            self.log("מתחיל ייבוא נתונים...", "info")
            seed, workbooks, errors = build_seed(Path(folder_text))
            self.progress["value"] = 70
            for workbook in workbooks:
                if not any(workbook.name in error for error in errors):
                    count = workbook_question_count(seed, workbook)
                    self.log(f"קובץ עובד: {workbook.name} ({count} שאלות)", "success")
            for error in errors:
                self.log(error, "error")

            output = Path(output_text)
            write_runner(seed, output)
            self.progress["value"] = 100
            question_count = sum(
                len(questions)
                for subject in seed.values()
                for workbook_seed in subject.values()
                for questions in workbook_seed.values()
            )
            self.log(f"הקובץ נשמר בנתיב: {output.resolve()}", "success")
            self.log(f"סיכום: {len(workbooks)} קובצי Excel, {question_count} שאלות", "success")
            self.open_output_folder(output)
        except Exception as exc:
            self.log(f"שגיאה בייבוא ושמירה: {exc}", "error")
        finally:
            self.set_import_enabled(True)

    def open_output_folder(self, output: Path) -> None:
        folder = output.resolve().parent
        try:
            if os.name == "nt":
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.run(["open", folder], check=True)
            else:
                subprocess.run(["xdg-open", folder], check=True)
            self.log(f"תיקיית היעד נפתחה: {folder}", "success")
        except Exception as exc:
            self.log(f"לא ניתן לפתוח את תיקיית היעד: {exc}", "error")


if __name__ == "__main__":
    SeedBuilderApp().mainloop()
