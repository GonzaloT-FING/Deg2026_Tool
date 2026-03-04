import tkinter as tk
import pathlib
from pipelines.eis_pip import export_folder
from tkinter import filedialog, ttk



PIPELINE_OPTIONS = {
    "EIS": ["Nyquist plot", "Bode plot", "I vs pt", "T vs pt", "Equivalent circuit fit"],
    "CV": ["I vs t", "Peak current", "Onset potential"],
    "PC": ["V vs I completo", "V vs I last", "V vs t", "I vs t", "T vs t", "dV/dI", "Step Stability"],
    "OCP": ["V vs t", "V final", "Drift", "DeltaV"],
    "Deg": ["V vs t", "I vs t", "T vs t", "dV/dt", "dV/dt final", "Trend", "Degradation rate"],
    "Análisis multiple": ["EIS", "CV", "PC", "OCP", "Deg"]
}


class GamryProtocolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Procesamiento de datos Gamry Protocol")
        self.root.geometry("650x270")

        # Repository root = folder where this gui file lives
        self.repo_dir = pathlib.Path(__file__).resolve().parent

        # Default folders relative to the repo
        self.default_input_dir = self.repo_dir / "data"
        self.default_output_dir = self.repo_dir / "outputs"

        # Ensure output folder exists
        self.default_output_dir.mkdir(parents=True, exist_ok=True)

        self.options_window = None
        self.selected_options = {}

        self.status_var = tk.StringVar()
        self.status_var.set("Listo. Seleccione carpetas y un pipeline.")

        self.build_main_window()

    def build_main_window(self):
        input_label = tk.Label(self.root, text="Carpeta de entrada:")
        input_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.input_entry = tk.Entry(self.root, width=60)
        self.input_entry.grid(row=0, column=1, padx=10, pady=10)
        self.input_entry.insert(0, str(self.default_input_dir))

        browsein_button = tk.Button(
            self.root,
            text="Buscar...",
            command=lambda: self.browse_button("bin")
        )
        browsein_button.grid(row=0, column=2, padx=10, pady=10)

        output_label = tk.Label(self.root, text="Carpeta de salida:")
        output_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.output_entry = tk.Entry(self.root, width=60)
        self.output_entry.grid(row=1, column=1, padx=10, pady=10)
        self.output_entry.insert(0, str(self.default_output_dir))

        browseout_button = tk.Button(
            self.root,
            text="Buscar...",
            command=lambda: self.browse_button("bout")
        )
        browseout_button.grid(row=1, column=2, padx=10, pady=10)

        pipeline_label = tk.Label(self.root, text="Analizar:")
        pipeline_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.pipeline_combo = ttk.Combobox(
            self.root,
            values=list(PIPELINE_OPTIONS.keys()),
            state="readonly",
            width=57
        )
        self.pipeline_combo.grid(row=2, column=1, padx=10, pady=10)
        self.pipeline_combo.current(0)

        next_button = tk.Button(self.root, text="Next", command=self.pipeline_selected)
        next_button.grid(row=3, column=2, padx=10, pady=15)

        self.status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            anchor="w",
            fg="blue"
        )
        self.status_label.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="we")

    def set_status(self, message):
        self.status_var.set(message)

    def browse_button(self, button_type):
        if button_type == "bin":
            title = "Select input folder"
            start_dir = self.input_entry.get().strip() or str(self.default_input_dir)
        else:
            title = "Select output folder"
            start_dir = self.output_entry.get().strip() or str(self.default_output_dir)

        folder = filedialog.askdirectory(
            title=title,
            initialdir=start_dir
        )

        if not folder:
            self.set_status("Selección cancelada.")
            return

        if button_type == "bin":
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, folder)
            self.set_status("Carpeta de entrada seleccionada.")
        else:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)
            self.set_status("Carpeta de salida seleccionada.")

    def pipeline_selected(self):
        input_dir = self.input_entry.get().strip()
        output_dir = self.output_entry.get().strip()
        selected_pipeline = self.pipeline_combo.get().strip()

        if not input_dir or not output_dir or not selected_pipeline:
            self.set_status("Complete todos los campos antes de continuar.")
            return

        print(f"Directorio de entrada: {input_dir}")
        print(f"Directorio de salida: {output_dir}")
        print(f"Pipeline: {selected_pipeline}")

        self.set_status(f"Pipeline seleccionado: {selected_pipeline}")
        self.open_pipeline_window(selected_pipeline)

    def open_pipeline_window(self, selected_pipeline):
        options = PIPELINE_OPTIONS.get(selected_pipeline, [])

        if self.options_window is not None and self.options_window.winfo_exists():
            self.options_window.destroy()

        self.options_window = tk.Toplevel(self.root)
        self.options_window.title(f"{selected_pipeline} properties")
        self.options_window.geometry("400x300")

        tk.Label(
            self.options_window,
            text=f"Select {selected_pipeline} options:"
        ).pack(pady=10)

        option_vars = []

        for option in options:
            var = tk.BooleanVar(value=False)
            chk = ttk.Checkbutton(self.options_window, text=option, variable=var)
            chk.pack(anchor="w", padx=20, pady=4)
            option_vars.append((option, var))

        def confirm():
            selected = [name for name, var in option_vars if var.get()]
            self.selected_options[selected_pipeline] = selected

            print(f"Selected {selected_pipeline} options: {selected}")

            if selected:
                self.set_status(f"{selected_pipeline}: {len(selected)} opción(es) seleccionada(s).")
            else:
                self.set_status(f"{selected_pipeline}: no se seleccionaron opciones.")

            # Run the EIS export only for the EIS pipeline
            if selected_pipeline == "EIS":
                try:
                    input_dir = pathlib.Path(self.input_entry.get().strip())
                    output_dir = pathlib.Path(self.output_entry.get().strip())

                    exported_files = export_folder(input_dir, output_dir, selected)

                    if exported_files:
                        self.set_status(
                            f"EIS exportado: {len(exported_files)} archivo(s) .xlsx creado(s). "
                            f"Se generaron también los gráficos seleccionados."
                        )
                    else:
                        self.set_status("No se encontraron archivos .DTA con 'EISPOT' en la carpeta de entrada.")

                except Exception as e:
                    import traceback
                    print(traceback.format_exc())
                    self.set_status(f"Error en EIS: {type(e).__name__}: {e}")

                self.options_window.destroy()
                self.options_window = None

        ttk.Button(
            self.options_window,
            text="Confirmar",
            command=confirm
        ).pack(pady=15)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = GamryProtocolApp()
    app.run()