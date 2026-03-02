import tkinter as tk
from tkinter import filedialog


def browse_input():
    folder = filedialog.askdirectory(title="Select input folder")
    if folder:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, folder)


def run_program():
    input_dir = input_entry.get()
    output_dir = output_entry.get()

    print(f"Input directory: {input_dir}")
    print(f"Output directory: {output_dir}")


root = tk.Tk()
root.title("Lab Data Tool")
root.geometry("650x180")

input_label = tk.Label(root, text="Input folder:")
input_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

input_entry = tk.Entry(root, width=60)
input_entry.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(root, text="Browse...", command=browse_input)
browse_button.grid(row=0, column=2, padx=10, pady=10)

output_label = tk.Label(root, text="Output folder:")
output_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

output_entry = tk.Entry(root, width=60)
output_entry.grid(row=1, column=1, padx=10, pady=10)

run_button = tk.Button(root, text="Run", command=run_program)
run_button.grid(row=2, column=1, pady=20)

root.mainloop()