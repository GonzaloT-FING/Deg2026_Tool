import tkinter as tk
from tkinter import filedialog, ttk


PIPELINE_OPTIONS = {
    "EIS": ["Nyquist plot", "Bode plot", "Equivalent circuit fit", "V", "I vs pt", "T vs t"],
    "CV": ["I vs t", "Peak current", "Onset potential"],
    "PC": ["V vs I completo", "V vs I last", "V vs t", "I vs t", "T vs t", "dV/dI", "Step Stability"],
    "OCP": ["V vs t", "V final", "Drift", "DeltaV"],
    "Deg": ["V vs t", "I vs t", "T vs t", "dV/dt", "dV/dt final", "Trend", "Degradation rate"],
    "Análisis multiple": ["EIS", "CV", "PC", "OCP", "Deg"]
}


def browse_button(button_type):
    title = "Select input folder" if button_type == "bin" else "Select output folder"
    folder = filedialog.askdirectory(title=title)

    if not folder:
        return

    if button_type == "bin":
        input_entry.delete(0, tk.END)
        input_entry.insert(0, folder)
    else:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder)


def pipeline_selected():
    input_dir = input_entry.get().strip()
    output_dir = output_entry.get().strip()
    selected_pipeline = pipeline_combo.get().strip()

    if not input_dir or not output_dir or not selected_pipeline:
        print("Complete todos los campos.")
        return

    print(f"Directorio de entrada: {input_dir}")
    print(f"Directorio de salida: {output_dir}")
    print(f"Pipeline: {selected_pipeline}")

    open_pipeline_window(selected_pipeline)


def open_pipeline_window(selected_pipeline):
    options = PIPELINE_OPTIONS.get(selected_pipeline, [])

    window = tk.Toplevel(root)
    window.title(f"{selected_pipeline} properties")
    window.geometry("400x300")

    tk.Label(window, text=f"Select {selected_pipeline} options:").pack(pady=10)

    option_vars = []

    for option in options:
        var = tk.BooleanVar(value=False)
        chk = ttk.Checkbutton(window, text=option, variable=var)
        chk.pack(anchor="w", padx=20, pady=4)
        option_vars.append((option, var))

    def confirm():
        selected = [name for name, var in option_vars if var.get()]
        print(f"Selected {selected_pipeline} options: {selected}")
        window.destroy()

    ttk.Button(window, text="Confirm", command=confirm).pack(pady=15)


root = tk.Tk()
root.title("Procesamiento de datos Gamry Protocol")
root.geometry("650x220")

input_label = tk.Label(root, text="Carpeta de entrada:")
input_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

input_entry = tk.Entry(root, width=60)
input_entry.grid(row=0, column=1, padx=10, pady=10)

browsein_button = tk.Button(root, text="Buscar...", command=lambda: browse_button("bin"))
browsein_button.grid(row=0, column=2, padx=10, pady=10)

output_label = tk.Label(root, text="Carpeta de salidar:")
output_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

output_entry = tk.Entry(root, width=60)
output_entry.grid(row=1, column=1, padx=10, pady=10)

browseout_button = tk.Button(root, text="Buscar...", command=lambda: browse_button("bout"))
browseout_button.grid(row=1, column=2, padx=10, pady=10)

pipeline_label = tk.Label(root, text="Analizar:")
pipeline_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

pipeline_combo = ttk.Combobox(
    root,
    values=list(PIPELINE_OPTIONS.keys()),
    state="readonly",
    width=57
)
pipeline_combo.grid(row=2, column=1, padx=10, pady=10)
pipeline_combo.current(0)

next_button = tk.Button(root, text="Next", command=pipeline_selected)
next_button.grid(row=3, column=2, padx=10, pady=20)

root.mainloop()