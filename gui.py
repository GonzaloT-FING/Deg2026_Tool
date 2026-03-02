import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import traceback

# Import your pipeline(s)
from pipelines.eis_pip import run as eis_run

# Pipeline registry: add more pipelines here later
PIPELINES = {
    "EIS (.DTA / EISPOT*)": {
        "func": eis_run,
        "input_kind": "folder",   # eis_pip expects a folder
    },
    # Example for later:
    # "Chronopotentiometry (.DTA)": {"func": cp_run, "input_kind": "folder"},
    # "Single file something": {"func": x_run, "input_kind": "file"},
}

PLOT_SIZES = ["small", "medium", "large"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Processing Tool")
        self.resizable(False, False)

        self.input_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value=str(Path.cwd() / "output"))
        self.pipeline_choice = tk.StringVar(value=list(PIPELINES.keys())[0])
        self.plot_size = tk.StringVar(value="medium")

        self._build_ui()

    def _build_ui(self):
        pad = 10
        frm = ttk.Frame(self, padding=pad)
        frm.grid(row=0, column=0, sticky="nsew")
        frm.columnconfigure(1, weight=1)

        # Input
        ttk.Label(frm, text="Input:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, width=52, textvariable=self.input_path).grid(row=0, column=1, padx=(6, 6), sticky="we")
        ttk.Button(frm, text="Browse…", command=self.pick_input).grid(row=0, column=2)

        # Output
        ttk.Label(frm, text="Output:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frm, width=52, textvariable=self.output_dir).grid(row=1, column=1, padx=(6, 6), pady=(8, 0), sticky="we")
        ttk.Button(frm, text="Browse…", command=self.pick_output).grid(row=1, column=2, pady=(8, 0))

        # Pipeline
        ttk.Label(frm, text="Pipeline:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        cb = ttk.Combobox(frm, textvariable=self.pipeline_choice, values=list(PIPELINES.keys()),
                          state="readonly", width=49)
        cb.grid(row=2, column=1, padx=(6, 6), pady=(8, 0), sticky="we")
        cb.bind("<<ComboboxSelected>>", lambda e: self._on_pipeline_changed())

        # Plot size (optional but useful)
        ttk.Label(frm, text="Plot size:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(frm, textvariable=self.plot_size, values=PLOT_SIZES,
                     state="readonly", width=49).grid(row=3, column=1, padx=(6, 6), pady=(8, 0), sticky="we")

        # Run button
        self.run_btn = ttk.Button(frm, text="Run", command=self.run_processing)
        self.run_btn.grid(row=4, column=0, columnspan=3, pady=(12, 0))

    def _on_pipeline_changed(self):
        """Optional: clear input when pipeline changes to avoid wrong type (file vs folder)."""
        self.input_path.set("")

    def pick_input(self):
        choice = self.pipeline_choice.get()
        input_kind = PIPELINES[choice]["input_kind"]

        if input_kind == "folder":
            p = filedialog.askdirectory(title="Select input folder")
        else:
            p = filedialog.askopenfilename(title="Select input file")

        if p:
            self.input_path.set(p)

    def pick_output(self):
        p = filedialog.askdirectory(title="Select output folder")
        if p:
            self.output_dir.set(p)

    def run_processing(self):
        choice = self.pipeline_choice.get()
        pipeline = PIPELINES[choice]["func"]
        input_kind = PIPELINES[choice]["input_kind"]

        in_str = self.input_path.get().strip()
        out_str = self.output_dir.get().strip()

        if not in_str:
            messagebox.showwarning("Missing input", "Please select an input path.")
            return

        input_path = Path(in_str)
        output_dir = Path(out_str)
        output_dir.mkdir(parents=True, exist_ok=True)

        # Validate input type
        if input_kind == "folder" and not input_path.is_dir():
            messagebox.showerror("Invalid input", "This pipeline expects a folder as input.")
            return
        if input_kind == "file" and not input_path.is_file():
            messagebox.showerror("Invalid input", "This pipeline expects a file as input.")
            return

        self.run_btn.config(state="disabled")

        def worker():
            try:
                # eis_pip signature: run(input_path, output_dir, plot_size="medium")
                pipeline(input_path, output_dir, plot_size=self.plot_size.get())
                self.after(0, lambda: messagebox.showinfo("Done", "Processing finished."))
            except Exception:
                err = traceback.format_exc()
                self.after(0, lambda: messagebox.showerror("Error", err))
            finally:
                self.after(0, lambda: self.run_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()