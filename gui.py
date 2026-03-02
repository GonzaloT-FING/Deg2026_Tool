import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import traceback
import os
import sys

# ---- Replace this with your real processing function ----
def process_data(input_path: Path, output_dir: Path, progress_cb=None, status_cb=None):
    """
    Your real pipeline should live elsewhere (e.g., processing.py) and be imported.
    Call progress_cb(0..100) if you want progress updates.
    """
    import time
    output_dir.mkdir(parents=True, exist_ok=True)

    if status_cb: status_cb("Starting processing...")
    for i in range(101):
        time.sleep(0.02)  # simulate work
        if progress_cb: progress_cb(i)
        if status_cb and i in (5, 30, 60, 90): status_cb(f"Working... {i}%")

    # simulate output file
    (output_dir / "result.txt").write_text(f"Processed: {input_path}\n", encoding="utf-8")
    if status_cb: status_cb("Done!")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Processing Tool")
        self.resizable(False, False)

        self.input_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value=str(Path.cwd() / "output"))
        self.status = tk.StringVar(value="Ready.")
        self.progress = tk.IntVar(value=0)

        self._build_ui()

    def _build_ui(self):
        pad = 10
        frm = ttk.Frame(self, padding=pad)
        frm.grid(row=0, column=0, sticky="nsew")

        # Input row
        ttk.Label(frm, text="Input:").grid(row=0, column=0, sticky="w")
        entry_in = ttk.Entry(frm, width=45, textvariable=self.input_path)
        entry_in.grid(row=0, column=1, padx=(6, 6), sticky="we")
        ttk.Button(frm, text="Browse…", command=self.pick_input).grid(row=0, column=2)

        # Output row
        ttk.Label(frm, text="Output:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        entry_out = ttk.Entry(frm, width=45, textvariable=self.output_dir)
        entry_out.grid(row=1, column=1, padx=(6, 6), pady=(8, 0), sticky="we")
        ttk.Button(frm, text="Browse…", command=self.pick_output).grid(row=1, column=2, pady=(8, 0))

        # Buttons row
        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=3, pady=(12, 0), sticky="we")

        self.run_btn = ttk.Button(btns, text="Run", command=self.run_processing)
        self.run_btn.pack(side="left")

        ttk.Button(btns, text="Open Output", command=self.open_output).pack(side="left", padx=8)
        ttk.Button(btns, text="Quit", command=self.destroy).pack(side="right")

        # Progress + status
        self.pb = ttk.Progressbar(frm, orient="horizontal", length=380, mode="determinate",
                                 variable=self.progress, maximum=100)
        self.pb.grid(row=3, column=0, columnspan=3, pady=(12, 0), sticky="we")

        ttk.Label(frm, textvariable=self.status).grid(row=4, column=0, columnspan=3, pady=(6, 0), sticky="w")

    def pick_input(self):
        p = filedialog.askopenfilename(title="Select input file")
        if p:
            self.input_path.set(p)

    def pick_output(self):
        p = filedialog.askdirectory(title="Select output folder")
        if p:
            self.output_dir.set(p)

    def set_status(self, text: str):
        self.status.set(text)

    def set_progress(self, value: int):
        self.progress.set(int(value))

    def run_processing(self):
        in_str = self.input_path.get().strip()
        out_str = self.output_dir.get().strip()

        if not in_str:
            messagebox.showwarning("Missing input", "Please select an input file.")
            return

        input_path = Path(in_str)
        output_dir = Path(out_str)

        self.run_btn.config(state="disabled")
        self.set_progress(0)
        self.set_status("Running...")

        def worker():
            try:
                process_data(
                    input_path=input_path,
                    output_dir=output_dir,
                    progress_cb=lambda v: self.after(0, self.set_progress, v),
                    status_cb=lambda s: self.after(0, self.set_status, s),
                )
                self.after(0, lambda: messagebox.showinfo("Success", "Processing finished."))
            except Exception:
                err = traceback.format_exc()
                self.after(0, lambda: messagebox.showerror("Error", err))
                self.after(0, lambda: self.set_status("Error. See details."))
            finally:
                self.after(0, lambda: self.run_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def open_output(self):
        out = Path(self.output_dir.get())
        out.mkdir(parents=True, exist_ok=True)

        # Cross-platform open folder
        if sys.platform.startswith("win"):
            os.startfile(out)  # type: ignore
        elif sys.platform == "darwin":
            os.system(f'open "{out}"')
        else:
            os.system(f'xdg-open "{out}"')


if __name__ == "__main__":
    App().mainloop()