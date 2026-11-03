from __future__ import annotations

import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, ttk

from seed_builder import build_seed, write_runner


class SeedBuilderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Seed Runner Builder")
        self.geometry("820x620")
        self.minsize(720, 560)
        self.configure(bg="#1e1e1e")

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
        style.configure("Card.TLabelframe", background="#252526", foreground="#cccccc", bordercolor="#3c3c3c")
        style.configure("Card.TLabelframe.Label", background="#252526", foreground="#cccccc", font=("Segoe UI", 10, "bold"))
        style.configure("VSCode.TEntry", fieldbackground="#3c3c3c", foreground="#ffffff", insertcolor="#ffffff", borderwidth=1)
        style.configure("VSCode.TButton", background="#3c3c3c", foreground="#ffffff", borderwidth=0, padding=(14, 8))
        style.map("VSCode.TButton", background=[("active", "#505050")])
        style.configure("Import.TButton", background="#007acc", foreground="#ffffff", borderwidth=0, padding=(22, 10), font=("Segoe UI", 11, "bold"))
        style.map("Import.TButton", background=[("active", "#1488d4"), ("disabled", "#094771")])
        style.configure("Progress.Horizontal.TProgressbar", background="#007acc", troughcolor="#3c3c3c", bordercolor="#3c3c3c", lightcolor="#007acc", darkcolor="#007acc")

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=24)
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 18))
        ttk.Label(header, text="SEED RUNNER", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="יצירת קובץ הרצה מנתוני תרגול מתוך Excel", style="Subtitle.TLabel").pack(anchor="w", pady=(4, 0))

        self._add_path_row(main, "תיקיית תרגול", self.folder_var, self.choose_folder)
        self._add_path_row(main, "קובץ יעד", self.output_var, self.choose_output)

        log_frame = ttk.LabelFrame(main, text="לוג ייבוא", style="Card.TLabelframe", padding=10)
        log_frame.pack(fill="both", expand=True, pady=(4, 12))
        self.status = tk.Text(
            log_frame, height=12, state="disabled", wrap="word", font=("Consolas", 10),
            bg="#1e1e1e", fg="#cccccc", insertbackground="#cccccc", relief="flat", padx=10, pady=10,
        )
        self.status.pack(fill="both", expand=True)
        self.status.tag_configure("info", foreground="#cccccc")
        self.status.tag_configure("success", foreground="#89d185")
        self.status.tag_configure("error", foreground="#f48771")
        self.status.tag_configure("warning", foreground="#dcdcaa")

        bottom = ttk.Frame(main)
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", style="Progress.Horizontal.TProgressbar")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 18))
        self.import_button = ttk.Button(bottom, text="יבא נתונים", style="Import.TButton", command=self.build)
        self.import_button.pack(side="right", ipadx=18, ipady=6)
        self.log("מוכן לייבוא. בחרו תיקיית תרגול וקובץ יעד.", "info")

    def _add_path_row(self, parent: ttk.Frame, label: str, variable: tk.StringVar, callback) -> None:
        frame = ttk.LabelFrame(parent, text=label, style="Card.TLabelframe", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        ttk.Entry(frame, textvariable=variable, style="VSCode.TEntry").pack(side="right", fill="x", expand=True, padx=(10, 0))
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
            self.import_button.configure(state="disabled")
            self.progress.start(12)
            self.log("מתחיל ייבוא נתונים...", "info")
            seed, workbooks, errors = build_seed(Path(folder_text))
            for workbook in workbooks:
                if not any(workbook.name in error for error in errors):
                    count = sum(
                        len(questions)
                        for subject in seed.values()
                        for files in subject.values()
                        if workbook.stem in files
                        for questions in files[workbook.stem].values()
                    ) if workbook.stem else 0
                    self.log(f"קובץ עובד: {workbook.name} ({count} שאלות)", "success")
            for error in errors:
                self.log(error, "error")

            output = Path(output_text)
            write_runner(seed, output)
            question_count = sum(
                len(questions)
                for subject in seed.values()
                for workbook_seed in subject.values()
                for questions in workbook_seed.values()
            )
            self.log(f"קובץ SEED_RUNNER נשמר בהצלחה: {output}", "success")
            self.log(f"סיכום: {len(workbooks)} קובצי Excel, {question_count} שאלות", "success")
        except Exception as exc:
            self.log(f"שגיאה בייבוא ושמירה: {exc}", "error")
        finally:
            self.progress.stop()
            self.import_button.configure(state="normal")


if __name__ == "__main__":
    SeedBuilderApp().mainloop()
