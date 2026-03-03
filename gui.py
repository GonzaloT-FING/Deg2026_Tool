import tkinter as tk
from tkinter import filedialog


def browse_button(button_type):
    folder = filedialog.askdirectory(title="Select input folder")
    if button_type == "bin":
        if folder:
            input_entry.delete(0, tk.END)
            input_entry.insert(0, folder)
    elif button_type == "bout":
        if folder:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, folder)


def run_program():
    input_dir = input_entry.get()
    output_dir = output_entry.get()

    print(f"Input directory: {input_dir}")
    print(f"Output directory: {output_dir}")


root = tk.Tk()
root.title("Procesamiento de datos Gamry Protocol")
root.geometry("650x180")

input_label = tk.Label(root, text="Input folder:")
input_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

input_entry = tk.Entry(root, width=60)
input_entry.grid(row=0, column=1, padx=10, pady=10)

browsein_button = tk.Button(root, text="Browse...", command=lambda: browse_button("bin"))
browsein_button.grid(row=0, column=2, padx=10, pady=10)

output_label = tk.Label(root, text="Output folder:")
output_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

output_entry = tk.Entry(root, width=60)
output_entry.grid(row=1, column=1, padx=10, pady=10)

browseout_button = tk.Button(root, text="Browse...", command=lambda: browse_button("bout"))
browseout_button.grid(row=1, column=2, padx=10, pady=10)

run_button = tk.Button(root, text="Next", command=run_program)
run_button.grid(row=2, column=2, pady=20)

root.mainloop()