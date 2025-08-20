import tkinter as tk
from tkinter import filedialog, messagebox
import os
import report_builder_v7  # replace with your actual module name

def run_builder():
    excel_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if not excel_file:
        return

    word_template = filedialog.askopenfilename(title="Select Word Template", filetypes=[("Word Files", "*.docx")])
    if not word_template:
        return

    output_folder = filedialog.askdirectory(title="Select Output Folder")
    if not output_folder:
        return

    output_filename = output_name_entry.get().strip()
    if not output_filename:
        messagebox.showwarning("Missing filename", "Please enter an output file name.")
        return

    # Ensure .docx extension
    if not output_filename.lower().endswith(".docx"):
        output_filename += ".docx"

    try:
        report_builder_v7.main_with_inputs(
            excel_path=excel_file,
            word_path=word_template,
            output_folder=output_folder,
            output_file_name=output_filename
        )
        messagebox.showinfo("Success", f"Saved as: {output_filename}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# --- GUI setup ---
root = tk.Tk()
root.title("GHG Report Builder")

tk.Label(root, text="Output File Name (without .docx):").pack()
output_name_entry = tk.Entry(root, width=40)
output_name_entry.pack(pady=5)

btn = tk.Button(root, text="Run Report Builder", command=run_builder, padx=20, pady=10)
btn.pack(pady=20)

root.mainloop()
