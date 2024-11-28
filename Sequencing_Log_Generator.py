import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
from datetime import datetime
import openpyxl

def load_bbid_data(file_path):
    print(f"Loading BBID data from: {file_path}")
    try:
        bbid_df = pd.read_excel(file_path, sheet_name='WGK - Initiated')
        print(f"BBID data loaded successfully with {len(bbid_df)} records")
        return bbid_df
    except Exception as e:
        print(f"Error loading BBID data: {e}")
        return None

def process_file(file_path, bbid_data):
    print(f"Processing file: {file_path}")
    
    try:
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, skiprows=5, engine='xlrd')
        else:
            df = pd.read_excel(file_path, skiprows=5)
        
        folder_column = None
        for col in df.columns:
            if df[col].apply(lambda x: isinstance(x, str) and '._.' in x).any():
                folder_column = col
                break
        
        if folder_column is None:
            print(f"No suitable column found in {file_path}")
            return None
        
        folder_names = []
        job_ids = []
        vector_ids = []
        for value in df[folder_column]:
            if isinstance(value, str) and '._.' in value:
                parts = value.split('._.')
                if len(parts) >= 2:
                    folder_name = parts[1] + '.' + parts[0]
                    folder_names.append(folder_name)
                    job_ids.append(parts[1])
                    vector_ids.append(parts[0])

        if not folder_names:
            print("No valid data extracted. Skipping file.")
            return None

        data = {'Folder name': folder_names, 'Job ID': job_ids, 'Vector ID': vector_ids}
        unique_data = pd.DataFrame(data).drop_duplicates()

        if 'JOB (WORK) ID' not in bbid_data.columns:
            print("Error: 'JOB (WORK) ID' column not found in BBID data")
            return None

        unique_data = unique_data.merge(bbid_data[['JOB (WORK) ID', 'BBID']], how='left', left_on='Job ID', right_on='JOB (WORK) ID')
        unique_data = unique_data.drop(columns=['JOB (WORK) ID'])

        log_df = unique_data.reindex(columns=['Folder name', 'Job ID', 'Vector ID', 'BBID', 'Sg SS OK', 'Sg DS OK', 'Sg Mutation or FAIL', 'Sg Primer to repeat'])
        
        seq_plate = os.path.splitext(os.path.basename(file_path))[0]
        log_df.insert(0, 'Seq Plate', seq_plate)
        
        print(f"Processed data for: {seq_plate}")
        return log_df
    
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return None

def run_script():
    source_dir = source_entry.get()
    log_dir = log_entry.get()
    bbid_source_file = r'Z:\Gene Synthesis\3.0 In-House Gene\3.5 Job Log\3.5 In-house Progress v3.XLSX'

    if not os.path.exists(source_dir):
        status_label.config(text=f"Source directory does not exist: {source_dir}")
        status_label.config(style="Red.TLabel")
        return
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        print(f"Created log directory: {log_dir}")

    bbid_data = load_bbid_data(bbid_source_file)
    if bbid_data is None:
        status_label.config(text="Failed to load BBID data, no log files created.")
        status_label.config(style="Red.TLabel")
        return

    current_date = datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"{current_date} - log.xlsx")
    
    all_data = []
    for file in os.listdir(source_dir):
        if file.endswith('.xls') or file.endswith('.xlsx'):
            file_path = os.path.join(source_dir, file)
            processed_data = process_file(file_path, bbid_data)
            if processed_data is not None:
                all_data.append(processed_data)
    
    if all_data:
        combined_data = pd.concat(all_data, ignore_index=True)
        combined_data.to_excel(log_file, sheet_name='Combined Log', index=False)
        status_label.config(text=f"Log file has been created successfully: {log_file}")
        status_label.config(style="Green.TLabel")
    else:
        status_label.config(text="No data processed. Log file not created.")
        status_label.config(style="Red.TLabel")

# Create the main window
window = tk.Tk()
window.title("Sequencing Log Processor")
window.geometry("500x300")

# Create styles for colored labels
style = ttk.Style()
style.configure("Red.TLabel", foreground="red")
style.configure("Green.TLabel", foreground="green")

# Create a main frame
main_frame = ttk.Frame(window, padding="10")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Configure grid
window.columnconfigure(0, weight=1)
window.rowconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

# Source Directory
ttk.Label(main_frame, text="Source Directory:").grid(column=0, row=0, sticky=tk.W, pady=5)
source_entry = ttk.Entry(main_frame, width=50)
source_entry.grid(column=1, row=0, sticky=(tk.W, tk.E), pady=5)
ttk.Button(main_frame, text="Browse", command=lambda: source_entry.delete(0, tk.END) or source_entry.insert(0, filedialog.askdirectory())).grid(column=2, row=0, sticky=tk.W, padx=5, pady=5)

# Log Directory
ttk.Label(main_frame, text="Log Directory:").grid(column=0, row=1, sticky=tk.W, pady=5)
log_entry = ttk.Entry(main_frame, width=50)
log_entry.grid(column=1, row=1, sticky=(tk.W, tk.E), pady=5)
ttk.Button(main_frame, text="Browse", command=lambda: log_entry.delete(0, tk.END) or log_entry.insert(0, filedialog.askdirectory())).grid(column=2, row=1, sticky=tk.W, padx=5, pady=5)

# Run Button
run_button = ttk.Button(main_frame, text="Run", command=run_script)
run_button.grid(column=1, row=2, pady=20)

# Status Label
status_label = ttk.Label(main_frame, text="", wraplength=480)
status_label.grid(column=0, row=3, columnspan=3, sticky=(tk.W, tk.E))

# Start the GUI event loop
window.mainloop()