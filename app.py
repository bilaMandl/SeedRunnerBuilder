from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from seed_builder import build_seed, write_runner


class SeedBuilderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Seed Runner Builder")
        self.geometry("720x430")
        self.resizable(False, False)

        self.folder_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / "SEED_RUNNER.py"))

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=24)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="בניית SEED_RUNNER.py", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 18))

        self._add_path_row(main, "תיקיית תרגול", self.folder_var, self.choose_folder)
        self._add_path_row(main, "קובץ יעד", self.output_var, self.choose_output)

        self.status = tk.Text(main, height=9, state="disabled", wrap="word", font=("Segoe UI", 10))
        self.status.pack(fill="both", expand=True, pady=(16, 12))

        ttk.Button(main, text="בנה ושמור", command=self.build).pack(side="right", ipadx=24, ipady=8)

    def _add_path_row(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        choose_callback,
    ) -> None:
        frame = ttk.LabelFrame(parent, text=label, padding=10)
        frame.pack(fill="x", pady=(0, 10))
        ttk.Entry(frame, textvariable=variable).pack(side="right", fill="x", expand=True, padx=(10, 0))
        ttk.Button(frame, text="בחירה", command=choose_callback).pack(side="right", padx=(10, 0))

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="בחר את תיקיית תרגול")
        if folder:
            self.folder_var.set(folder)

    def choose_output(self) -> None:
        output = filedialog.asksaveasfilename(
            title="שמירת SEED_RUNNER.py",
            defaultextension=".py",
            initialfile="SEED_RUNNER.py",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
        )
        if output:
            self.output_var.set(output)

    def set_status(self, text: str) -> None:
        self.status.configure(state="normal")
        self.status.delete("1.0", "end")
        self.status.insert("end", text)
        self.status.configure(state="disabled")

    def build(self) -> None:
        folder_text = self.folder_var.get().strip()
        output_text = self.output_var.get().strip()
        if not folder_text:
            messagebox.showerror("חסר נתיב", "יש לבחור תיקיית תרגול.")
            return
        if not output_text:
            messagebox.showerror("חסר יעד", "יש לבחור נתיב לשמירת הקובץ.")
            return

        try:
            seed, workbooks, errors = build_seed(Path(folder_text))
            output = Path(output_text)
            write_runner(seed, output)
        except Exception as exc:
            self.set_status(f"שגיאה: {exc}")
            messagebox.showerror("הבנייה נכשלה", str(exc))
            return

        question_count = sum(
            len(questions)
            for subject in seed.values()
            for workbook in subject.values()
            for questions in workbook.values()
        )
        status = (
            f"הקובץ נוצר בהצלחה:\n{output}\n\n"
            f"קובצי Excel: {len(workbooks)}\nשאלות: {question_count}\n"
        )
        if errors:
            status += "\nקבצים שנכשלו:\n" + "\n".join(errors)
        self.set_status(status)
        messagebox.showinfo("הושלם", "SEED_RUNNER.py נשמר בהצלחה.")


if __name__ == "__main__":
    SeedBuilderApp().mainloop()
