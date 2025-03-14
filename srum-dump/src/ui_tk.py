import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
import pathlib
import os
import ctypes
import tempfile
import subprocess
import urllib.request


def show_live_system_warning():
    """Warn the user when they try to analyze the srum on their own live system."""
    def download_fget():
        webbrowser.open("https://github.com/MarkBaggett/srum-dump/blob/master/FGET.exe")

    def auto_extract():
        result = extract_live_file()
        if result:
            srum_path_entry.delete(0, tk.END)
            srum_path_entry.insert(0, result[0])
            reg_path_entry.delete(0, tk.END)
            reg_path_entry.insert(0, result[1])
        warning_window.destroy()

    warning_window = tk.Tk()
    warning_window.title("WARNING")
    warning_window.geometry("500x200")

    tk.Label(warning_window, text="It appears you're trying to open SRUDB.DAT from a live system.").pack()
    tk.Label(warning_window, text="Copying or reading that file while it is locked is unlikely to succeed.").pack()
    tk.Label(warning_window, text="First, use a tool such as FGET that can copy files that are in use.").pack()
    tk.Label(warning_window, text=r"Try: 'fget -extract c:\windows\system32\sru\srudb.dat <a destination path>'").pack()

    button_frame = tk.Frame(warning_window)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="Close", command=warning_window.destroy).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Download FGET", command=download_fget).pack(side=tk.LEFT, padx=5)
    if ctypes.windll.shell32.IsUserAnAdmin() == 1:
        tk.Button(button_frame, text="Auto Extract", command=auto_extract).pack(side=tk.LEFT, padx=5)

    warning_window.mainloop()


def get_user_input(options):
    srum_path = ""
    if os.path.exists("SRUDB.DAT"):
        srum_path = os.path.join(os.getcwd(), "SRUDB.DAT")
    temp_path = pathlib.Path.cwd() / "SRUM_TEMPLATE2.XLSX"
    if temp_path.exists():
        temp_path = str(temp_path)
    else:
        temp_path = ""
    reg_path = ""
    if os.path.exists("SOFTWARE"):
        reg_path = os.path.join(os.getcwd(), "SOFTWARE")

    def browse_file(entry):
        file_path = filedialog.askopenfilename()
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def browse_folder(entry):
        folder_path = filedialog.askdirectory()
        if folder_path:
            entry.delete(0, tk.END)
            entry.insert(0, folder_path)

    def on_support_click(event):
        webbrowser.open("https://twitter.com/MarkBaggett")

    def on_ok():
        srum_path = srum_path_entry.get()
        out_dir = out_dir_entry.get()
        tem_path = tem_path_entry.get()
        reg_path = reg_path_entry.get()

        if not pathlib.Path(srum_path).exists() or not pathlib.Path(srum_path).is_file():
            messagebox.showerror("Error", "SRUM DATABASE NOT FOUND.")
            return
        if not os.path.exists(out_dir):
            messagebox.showerror("Error", "OUTPUT DIR NOT FOUND.")
            return
        if not pathlib.Path(tem_path).exists() or not pathlib.Path(tem_path).is_file():
            messagebox.showerror("Error", "SRUM TEMPLATE NOT FOUND.")
            return
        if reg_path and not pathlib.Path(reg_path).exists() and not pathlib.Path(reg_path).is_file():
            messagebox.showerror("Error", "REGISTRY File not found. (Leave field empty for None.)")
            return

        options.SRUM_INFILE = srum_path
        options.XLSX_OUTFILE = os.path.join(out_dir, "SRUM_DUMP_OUTPUT.xlsx")
        options.XLSX_TEMPLATE = tem_path
        options.reghive = reg_path if reg_path != "." else ""
        root.destroy()

    root = tk.Tk()
    root.title("SRUM_DUMP 2.6")
    root.geometry("600x400")

    tk.Label(root, text='REQUIRED: Path to SRUDB.DAT').pack(pady=5)
    srum_path_entry = tk.Entry(root, width=80)
    srum_path_entry.pack(pady=5)
    srum_path_entry.insert(0, srum_path)
    tk.Button(root, text="Browse", command=lambda: browse_file(srum_path_entry)).pack(pady=5)

    tk.Label(root, text='REQUIRED: Output folder for SRUM_DUMP_OUTPUT.xlsx').pack(pady=5)
    out_dir_entry = tk.Entry(root, width=80)
    out_dir_entry.pack(pady=5)
    out_dir_entry.insert(0, os.getcwd())
    tk.Button(root, text="Browse", command=lambda: browse_folder(out_dir_entry)).pack(pady=5)

    tk.Label(root, text='REQUIRED: Path to SRUM_DUMP Template').pack(pady=5)
    tem_path_entry = tk.Entry(root, width=80)
    tem_path_entry.pack(pady=5)
    tem_path_entry.insert(0, temp_path)
    tk.Button(root, text="Browse", command=lambda: browse_file(tem_path_entry)).pack(pady=5)

    tk.Label(root, text='RECOMMENDED: Path to registry SOFTWARE hive').pack(pady=5)
    reg_path_entry = tk.Entry(root, width=80)
    reg_path_entry.pack(pady=5)
    reg_path_entry.insert(0, reg_path)
    tk.Button(root, text="Browse", command=lambda: browse_file(reg_path_entry)).pack(pady=5)

    support_label = tk.Label(root, text="Click here for support via Twitter @MarkBaggett", fg="blue", cursor="hand2")
    support_label.pack(pady=5)
    support_label.bind("<Button-1>", on_support_click)

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Cancel", command=root.destroy).pack(side=tk.LEFT, padx=5)

    root.mainloop()
