import sys
import os
import tkinter as tk
import threading
import pandas as pd
import numpy as np
import time
import shutil
import tempfile

from openpyxl import load_workbook
from asammdf import MDF, Signal
from tkinter import Checkbutton, ttk, filedialog, messagebox


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    # PyInstaller creates a temp folder and stores path in _MEIPASS
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


jobs = []

def timing(func):
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        print(f"[TIMER] {func.__name__} took {end - start:.3f} seconds")
        return result
    return wrapper

# Asks user to select files for input, returns a tuple of (mf4_path, xlsm_path)
def add_job():
    mf4_path = filedialog.askopenfilename(
            title="Select MF4 File (Cancel to create new from Excel)",
            filetypes=[("MF4 files", "*.mf4"), ("All files", "*.*")]
        )
    if mf4_path:
        xlsm_paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel Macro-enabled files", "*.xlsm"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        jobs.append((mf4_path, list(xlsm_paths)))
    else:
        xlsm_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Macro-enabled files", "*.xlsm"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        jobs.append((None, [xlsm_path]))
    update_jobs_label()

# Updates the jobs readout to reflect current tasks
def update_jobs_label():
    update_text = ""
    for i, (mf4_path, xlsm_paths) in enumerate(jobs):
        update_text += f"Job {i + 1}:\n"
        if mf4_path:
            update_text += f"   MF4: {os.path.basename(mf4_path)}\n"
        for idx, xlsm_path in enumerate(xlsm_paths):
            update_text += f"   Excel {idx+1}: {os.path.basename(xlsm_path)}\n"
    jobs_label.config(text=update_text)

@timing
def read_xlsm_for_merge(xlsm_path, mdf_orig=None):
    # Check available sheets
    xl = pd.ExcelFile(xlsm_path, engine='openpyxl')
    if 'Uniplot' in xl.sheet_names:
        sheet_to_read = 'Uniplot'
    elif '10 Hz Data' in xl.sheet_names:
        sheet_to_read = '10 Hz Data'
    else:
        raise ValueError("Neither 'Uniplot' nor '10 Hz Data' sheet found in the Excel file.")

    # Step 1: Read only the header
    header_df = pd.read_excel(xlsm_path, sheet_name=sheet_to_read, nrows=0, engine='openpyxl')
    excel_columns = list(header_df.columns)

    # Step 2: Compare with MF4 channels if provided 
    if mdf_orig is not None:
        mf4_channels = set(mdf_orig.channels_db.keys())
        needed_columns = [col for col in excel_columns if col not in mf4_channels or col == 'Test Time']
    else:
        needed_columns = excel_columns

    # Step 3: Read only needed columns
    df = pd.read_excel(xlsm_path, sheet_name=sheet_to_read, usecols=needed_columns, engine='openpyxl')
    variable_names = [str(v) if pd.notnull(v) else '' for v in list(df.columns)]
    units = [str(u) if pd.notnull(u) else '' for u in df.iloc[0].tolist()]
    df_data = df.iloc[1:].reset_index(drop=True)
    df_data.columns = variable_names

    if 'Test Time' not in variable_names:
        raise ValueError(f"The '{sheet_to_read}' sheet must contain a 'Test Time' column.")

    # Convert all columns except 'Test Time' to numeric
    for col in df_data.columns:
        if col != 'Test Time':
            df_data[col] = pd.to_numeric(df_data[col], errors='coerce')

    # Remove columns that are all NaN (except 'Test Time')
    keep_cols = []
    keep_units = []
    for idx, col in enumerate(variable_names):
        if col == 'Test Time':
            keep_cols.append(col)
            keep_units.append(units[idx])
        elif col in df_data.columns and df_data[col].notna().any():
            keep_cols.append(col)
            keep_units.append(units[idx])
    df_data = df_data[keep_cols]
    print(f'{sheet_to_read} read ', df_data.shape)
    return df_data, keep_cols, keep_units

# Reads the xlsm file and extracts the 'Cumulatives' sheet, saving it as a new Excel file; throws error if the sheet is missing
@timing
def CRS_output(xlsm_paths):
    for sheet in xlsm_paths:
        wb = load_workbook(sheet, read_only=True, data_only=True)
        if 'Cumulatives' in wb.sheetnames:
            ws = wb['Cumulatives']
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(list(row))
            df = pd.DataFrame(data)
        else:
            raise ValueError("The Excel file is missing a 'Cumulatives' sheet, no CRS will be given.")
        output_path = os.path.splitext(sheet)[0] + '_summary.xlsx'
        df.to_excel(output_path, index=False, header=False, engine='openpyxl')

@timing
def merge_xlsm_to_mf4(mf4_path, xlsm_paths, progress_callback=None):
    """
    xlsm_paths: list of Excel file paths
    """
    mdf_orig = MDF(mf4_path)
    engspd_channels = [ch for ch in mdf_orig.channels_db if ch.lower().startswith('engine_speed')]
    if not engspd_channels:
        raise ValueError("No engine_speed channel found in MF4.")

    ch_name = engspd_channels[0]
    engspd = mdf_orig.get(ch_name)

    mdf = MDF()
    excel_signals_total = 0

    # Add all MF4 channels first (unaligned, will be trimmed later if needed)
    for name in mdf_orig.channels_db:
        for group, index in mdf_orig.channels_db[name]:
            sig = mdf_orig.get(name, group=group, index=index)
            mdf.append(sig)

    for xlsm_path in xlsm_paths:
        df, variable_names, units = read_xlsm_for_merge(xlsm_path)
        excel_speed_col = next((name for name in variable_names if name.lower() in ['engine speed', 'engine_speed']), None)
        if not excel_speed_col or excel_speed_col not in df.columns:
            continue

        # Get MF4 engine speed and time
        mf4_speed = engspd.samples
        mf4_time = engspd.timestamps

        # Get Excel engine speed and time
        excel_speed = df[excel_speed_col].to_numpy(dtype=float)
        excel_time = df['Test Time'].to_numpy(dtype=float)

        # Normalize for cross-correlation
        excel_speed_norm = (excel_speed - np.mean(excel_speed)) / np.std(excel_speed)
        mf4_speed_norm = (mf4_speed - np.mean(mf4_speed)) / np.std(mf4_speed)

        # Cross-correlate (allow for any overlap)
        correlation = np.correlate(mf4_speed_norm, excel_speed_norm, mode='full')
        lags = np.arange(-len(excel_speed_norm) + 1, len(mf4_speed_norm))

        # Only consider lags where the aligned Excel time overlaps with MF4 time
        valid_lags = []
        for i, lag in enumerate(lags):
            # Calculate the time offset for this lag
            if lag >= 0:
                mf4_time_at_lag = mf4_time[lag]
                excel_time_at_zero = excel_time[0]
                time_offset = excel_time_at_zero - mf4_time_at_lag
                # Excel shifted left, must not start before MF4
                if (excel_time[0] - time_offset) < mf4_time[0]:
                    continue
                # Excel shifted right, must not end after MF4
                if (excel_time[-1] - time_offset) > mf4_time[-1]:
                    continue
            else:
                mf4_time_at_lag = mf4_time[0]
                excel_time_at_zero = excel_time[-lag]
                time_offset = excel_time_at_zero - mf4_time_at_lag
                # Excel shifted left, must not start before MF4
                if (excel_time[0] - time_offset) < mf4_time[0]:
                    continue
                # Excel shifted right, must not end after MF4
                if (excel_time[-1] - time_offset) > mf4_time[-1]:
                    continue
            # Overlap check passed
            valid_lags.append(i)

        if not valid_lags:
            raise ValueError("No sufficient overlap between MF4 and Excel engine speed signals.")

        # Find the best lag within valid lags
        best_idx = valid_lags[np.argmax(correlation[valid_lags])]
        lag = lags[best_idx]

        # Calculate alignment
        if lag >= 0:
            mf4_time_at_lag = mf4_time[lag]
            excel_time_at_zero = excel_time[0]
            time_offset = excel_time_at_zero - mf4_time_at_lag
        else:
            mf4_time_at_lag = mf4_time[0]
            excel_time_at_zero = excel_time[-lag]
            time_offset = excel_time_at_zero - mf4_time_at_lag

        # Now, align Excel signals to MF4 time base
        for idx, col in enumerate(variable_names):
            if col == 'Test Time' or col.strip() == '':
                continue
            if col in df.columns:
                samples = df[col].values.astype(float)
                if not np.isnan(samples).all() and len(samples) > 0:
                    unit = units[idx] if idx < len(units) else ''
                    signal = Signal(
                        samples=samples,
                        timestamps=excel_time - time_offset,  # align to MF4
                        name=col,
                        unit=unit,
                        comment=''
                    )
                    mdf.append(signal)
                    excel_signals_total += 1

    folder = os.path.dirname(xlsm_paths[0])
    base = os.path.splitext(os.path.basename(xlsm_paths[0]))[0]
    output_path = os.path.join(folder, base + "_merged.mf4")
    mdf.save(output_path, compression=1)
    return output_path, excel_signals_total

def remove_job():
    if jobs:
        jobs.pop()
        update_jobs_label()
    else:
        messagebox.showinfo("Info", "No tasks to undo.")

# Tkinter GUI Handling
def threaded_merge(jobs):
    successes = []
    errors = []

    def handle_success(output_path, added):
        successes.append((output_path, added))

    def handle_error(e, xlsm_path):
        errors.append((e, xlsm_path))

    def prepare_progress():
        task_btn_frame.pack_forget()
        convert_btn.pack_forget()
        progress.pack()
        progress.start(10)
    
    def show_results():
        progress.stop()
        progress.pack_forget()
        task_btn_frame.pack(pady=10)
        convert_btn.pack(pady=10)
        results = ""
        if successes:
            results = "Completed Tasks:\n"
            for output_path, added in successes:
                results += f"Added {added} signals to {output_path}\n"

        if errors:
            results += "\nFailed Tasks:\n"
            for error, xlsm_path in errors:
                # ...existing code...
                if isinstance(xlsm_path, (list, tuple)):
                    excel_files = ", ".join(os.path.basename(x) for x in xlsm_path)
                else:
                    excel_files = os.path.basename(xlsm_path)
                results += f"Error: {error}\nExcel File: {excel_files}\n"
                # ...existing code...
        messagebox.showinfo("Results", results)
        jobs_label.config(text="No tasks queued")

    @timing
    def run_merge():
        jobs_copy = jobs[:]
        for mf4_path, xlsm_paths in jobs_copy:
            try:
                output_path, added = merge_xlsm_to_mf4(mf4_path, xlsm_paths)
                try:
                    CRS_output(xlsm_paths) # !TODO allow for CRS output for multiple Excel inputs
                except Exception as e:
                    handle_error(e, xlsm_paths)
                handle_success(output_path, added)
            except Exception as e:
                handle_error(e, xlsm_paths)
            jobs.pop(0)
            root.after(0, update_jobs_label)
        root.after(0, show_results)

    root.after(0, prepare_progress)
    threading.Thread(target=run_merge, daemon=True).start()


root = tk.Tk()
root.title("MF4 + Excel Merge and Convert")
root.geometry("400x600")
root.resizable(False, True)

icon_path = resource_path(os.path.join("icons", "dumarey_favicon.ico"))

# Fix for PyInstaller: extract icon to temp file if running as executable
if hasattr(sys, '_MEIPASS'):
    temp_icon = tempfile.NamedTemporaryFile(delete=False, suffix='.ico')
    shutil.copyfile(icon_path, temp_icon.name)
    root.iconbitmap(temp_icon.name)
else:
    root.iconbitmap(icon_path)

style = ttk.Style(root)
style.theme_use('default')
green = "#28a745"
style.configure("Green.TButton", foreground="white", background=green)
style.map("Green.TButton",
        background=[('active', '#218838'), ('!active', green)])
red = "#972e26"
style.configure("Red.TButton", foreground="white", background=red)
style.map("Red.TButton",
        background=[('active', '#c82333'), ('!active', red)])

main_frame = ttk.Frame(root, padding=20, style="Main.TFrame")
main_frame.pack(expand=True, fill='both')

title_label = ttk.Label(main_frame, text="MF4/Excel Merge and Convert", font=("Segoe UI", 16, "bold"))
title_label.pack(pady=(0, 10))

desc_label = ttk.Label(
    main_frame,
    text="Select an MF4 and an Excel file to merge.\n"
        "Or cancel MF4 selection to convert Excel to MF4.",
    font=("Segoe UI", 10),
    style="Main.TLabel"
)
desc_label.pack(pady=(0, 20))

label_frame = ttk.LabelFrame(main_frame, text="Queued Tasks")
label_frame.pack(fill='both', expand=True, pady=(0, 10), padx=20)

jobs_label = ttk.Label(label_frame, text="No tasks queued", font=("Segoe UI", 9), foreground="gray")
jobs_label.pack(pady=(5, 0))

task_btn_frame = ttk.Frame(main_frame)
task_btn_frame.pack(pady=10)

add_task_btn = ttk.Button(task_btn_frame, text="Add Task", command=lambda: add_job())
add_task_btn.configure(width=16)
add_task_btn.pack(side='left', padx=(0, 5))

undo_task_btn = ttk.Button(task_btn_frame, text="Undo", command=lambda: remove_job())
undo_task_btn.configure(style="Red.TButton")
undo_task_btn.pack(side='left')

convert_btn = ttk.Button(main_frame, text="Merge and Convert", command=lambda: threaded_merge(jobs), style="Green.TButton")
convert_btn.pack(pady=10)

progress = ttk.Progressbar(main_frame, mode='indeterminate', length=250)
progress.pack(pady=10)
progress.pack_forget()

status_var = tk.StringVar()
status_label = ttk.Label(main_frame, textvariable=status_var, font=("Segoe UI", 9), foreground="gray")
status_label.pack(pady=(5, 0))

root.mainloop()