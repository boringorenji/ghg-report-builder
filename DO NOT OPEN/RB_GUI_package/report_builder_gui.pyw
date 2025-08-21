import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import os
import json
import traceback
import sys  # NEW: for resource_path

# IMPORTANT: your backend module that exposes main_with_inputs(...)
import report_builder

APP_TITLE = "GHG Report Builder"
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".ghg_report_builder_gui.json")

# ---------- Helpers for packaging (works in dev & PyInstaller) ----------

def resource_path(relpath: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relpath)
    return os.path.join(os.path.abspath("."), relpath)

def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_settings(d):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ---------- Worker logic ----------

def set_running(btns, running: bool, note_var: tk.StringVar, note: str = ""):
    for b in btns:
        try:
            b.config(state=("disabled" if running else "normal"))
        except Exception:
            pass
    note_var.set(note)

def run_builder_worker(excel_file, word_template, output_folder, output_filename, btns, note_var):
    try:
        set_running(btns, True, note_var, "Working… this may take a minute.")

        # Ensure .docx extension
        if not output_filename.lower().endswith(".docx"):
            output_filename += ".docx"

        # Ensure output folder exists
        os.makedirs(output_folder, exist_ok=True)

        report_builder.main_with_inputs(
            excel_path=excel_file,
            word_path=word_template,
            output_folder=output_folder,
            output_file_name=output_filename
        )

        save_settings({
            "excel_file": excel_file,
            "word_template": word_template,
            "output_folder": output_folder,
            "output_filename": output_filename,
        })

        set_running(btns, False, note_var, "Done!")
        messagebox.showinfo("Success", f"Saved as:\n{os.path.join(output_folder, output_filename)}")
    except Exception as e:
        set_running(btns, False, note_var, "")
        tb = traceback.format_exc(limit=400)
        messagebox.showerror("Error", f"An error occurred:\n{e}\n\nDetails:\n{tb}")

# ---------- Validators ----------

def validate_common(excel_file, word_template, output_folder, output_filename) -> bool:
    if not excel_file:
        messagebox.showwarning("Missing Excel", "Please select an Excel file (.xlsx).")
        return False
    if not word_template:
        messagebox.showwarning("Missing Template", "Please select a Word template (.docx).")
        return False
    if not output_folder:
        messagebox.showwarning("Missing Folder", "Please choose an output folder.")
        return False
    if not output_filename:
        messagebox.showwarning("Missing Filename", "Please enter an output file name.")
        return False
    return True

# ---------- Browse helper ----------

def browse_into(var: tk.StringVar, kind: str):
    path = None
    if kind == "excel":
        path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
    elif kind == "word":
        path = filedialog.askopenfilename(title="Select Word Template", filetypes=[("Word Files", "*.docx")])
    elif kind == "dir":
        path = filedialog.askdirectory(title="Select Output Folder")
    if path:
        var.set(path)

def run_advanced(btns, note_var, excel_var, word_var, outdir_var, name_var):
    excel_file = excel_var.get().strip()
    word_template = word_var.get().strip()
    output_folder = outdir_var.get().strip()
    output_filename = name_var.get().strip()

    if not validate_common(excel_file, word_template, output_folder, output_filename):
        return

    t = threading.Thread(
        target=run_builder_worker,
        args=(excel_file, word_template, output_folder, output_filename, btns, note_var),
        daemon=True,
    )
    t.start()
    set_running(btns, True, note_var, "Starting…")

# ---------- Build GUI ----------

def build_gui():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("680x400")

    # --- NEW: set window + taskbar icon ---
    try:
        root.iconbitmap(resource_path(os.path.join("icon", "ghg-rep-builder.ico")))  # Title bar / Alt-Tab on Windows
    except Exception:
        pass

    # Optional: if a PNG exists, also set iconphoto (helps taskbar on some Tk builds)
    png_fallback = resource_path(os.path.join("icon", "ghg-rep-builder.png"))
    if os.path.exists(png_fallback):
        try:
            icon_img = tk.PhotoImage(file=png_fallback)
            root.iconphoto(True, icon_img)
            root._icon_img_ref = icon_img  # prevent garbage collection
        except Exception:
            pass

    # Optional: unique AppID so Windows taskbar groups under your icon/name
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GHG.Report.Builder")
    except Exception:
        pass
    # --- end NEW ---

    settings = load_settings()
    status_var = tk.StringVar(value="")

    # Single tab (Advanced only)
    tab_adv = ttk.Frame(root)
    tab_adv.pack(fill="both", expand=True)

    adv_padx, adv_pady = 12, 8

    excel_var = tk.StringVar(value=settings.get("excel_file", ""))
    word_var = tk.StringVar(value=settings.get("word_template", ""))
    outdir_var = tk.StringVar(value=settings.get("output_folder", ""))
    name_var = tk.StringVar(value=settings.get("output_filename", "output.docx").replace(".docx", ""))

    # Excel row
    frm1 = ttk.Frame(tab_adv); frm1.pack(fill="x", padx=adv_padx, pady=adv_pady)
    ttk.Label(frm1, text="Excel (.xlsx):", width=20).pack(side="left")
    ttk.Entry(frm1, textvariable=excel_var).pack(side="left", fill="x", expand=True)
    ttk.Button(frm1, text="Browse…", command=lambda: browse_into(excel_var, "excel")).pack(side="left", padx=6)

    # Word row
    frm2 = ttk.Frame(tab_adv); frm2.pack(fill="x", padx=adv_padx, pady=adv_pady)
    ttk.Label(frm2, text="Word template (.docx):", width=20).pack(side="left")
    ttk.Entry(frm2, textvariable=word_var).pack(side="left", fill="x", expand=True)
    ttk.Button(frm2, text="Browse…", command=lambda: browse_into(word_var, "word")).pack(side="left", padx=6)

    # Output folder row
    frm3 = ttk.Frame(tab_adv); frm3.pack(fill="x", padx=adv_padx, pady=adv_pady)
    ttk.Label(frm3, text="Output folder:", width=20).pack(side="left")
    ttk.Entry(frm3, textvariable=outdir_var).pack(side="left", fill="x", expand=True)
    ttk.Button(frm3, text="Browse…", command=lambda: browse_into(outdir_var, "dir")).pack(side="left", padx=6)

    # Output filename row
    frm4 = ttk.Frame(tab_adv); frm4.pack(fill="x", padx=adv_padx, pady=adv_pady)
    ttk.Label(frm4, text="Output file name:", width=20).pack(side="left")
    ttk.Entry(frm4, textvariable=name_var).pack(side="left", fill="x", expand=True)

    # Run button
    btn_adv_run = ttk.Button(tab_adv, text="Run Report Builder")
    btn_adv_run.pack(pady=14)

    ttk.Separator(tab_adv, orient="horizontal").pack(fill="x", padx=adv_padx, pady=8)
    ttk.Label(tab_adv, textvariable=status_var, foreground="#666").pack(pady=(0, 6))

    btns_adv = [btn_adv_run]
    btn_adv_run.config(command=lambda: run_advanced(btns_adv, status_var, excel_var, word_var, outdir_var, name_var))

    root.mainloop()

if __name__ == "__main__":
    build_gui()
