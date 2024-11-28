import os
import shutil
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Function to extract Job ID and Plasmid Number from filename
def extract_info(filename):
    pattern = re.compile(r'(.+?)_(.+?)_(.+)')
    match = pattern.match(filename)
    if match:
        plasmid_number = match.group(1)
        job_id = match.group(2).strip('.')
        return job_id, plasmid_number
    return None, None

# Function to get the BBID for a given job ID
def get_bbid_mapping(excel_path):
    if not os.path.exists(excel_path):
        print(f"Excel file does not exist: {excel_path}")
        return {}
    try:
        df = pd.read_excel(excel_path, sheet_name='WGK - Initiated')
        bbid_mapping = pd.Series(df['BBID'].values, index=df['JOB (WORK) ID']).to_dict()
        return bbid_mapping
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return {}

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("File Organizer and Reference Distributor")
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both')
        
        # Create tabs
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Organize Sequencing Files")
        self.notebook.add(self.tab2, text="Distribute Reference Files")
        
        # Build the GUI for each tab
        self.build_tab1()
        self.build_tab2()
    
    def build_tab1(self):
        # Script 1: Organize Sequencing Files
        # Labels and Entry fields for source directory and destination directory
        self.source_dir_label1 = ttk.Label(self.tab1, text="Sequencing files Directory:")
        self.source_dir_label1.grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.source_dir_entry1 = ttk.Entry(self.tab1, width=50)
        self.source_dir_entry1.grid(row=0, column=1, padx=5, pady=5)
        self.source_dir_button1 = ttk.Button(self.tab1, text="Browse", command=self.browse_source_dir1)
        self.source_dir_button1.grid(row=0, column=2, padx=5, pady=5)

        self.dest_dir_label1 = ttk.Label(self.tab1, text="Sorted sequencing folder Directory:")
        self.dest_dir_label1.grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.dest_dir_entry1 = ttk.Entry(self.tab1, width=50)
        self.dest_dir_entry1.grid(row=1, column=1, padx=5, pady=5)
        self.dest_dir_button1 = ttk.Button(self.tab1, text="Browse", command=self.browse_dest_dir1)
        self.dest_dir_button1.grid(row=1, column=2, padx=5, pady=5)

        # Start button
        self.start_button1 = ttk.Button(self.tab1, text="Start Process", command=self.start_process1)
        self.start_button1.grid(row=2, column=1, pady=20)
        
        # Status messages
        self.status_label1 = ttk.Label(self.tab1, text="", foreground="green")
        self.status_label1.grid(row=3, column=0, columnspan=3, padx=5, pady=5)
    
    def build_tab2(self):
        # Script 2: Distribute Reference Files
        # Labels and Entry fields for destination directory
        self.dest_base_dir_label2 = ttk.Label(self.tab2, text="Destination Directory:")
        self.dest_base_dir_label2.grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.dest_base_dir_entry2 = ttk.Entry(self.tab2, width=50)
        self.dest_base_dir_entry2.grid(row=0, column=1, padx=5, pady=5)
        self.dest_base_dir_button2 = ttk.Button(self.tab2, text="Browse", command=self.browse_dest_base_dir2)
        self.dest_base_dir_button2.grid(row=0, column=2, padx=5, pady=5)

        # Start button
        self.start_button2 = ttk.Button(self.tab2, text="Start Process", command=self.start_process2)
        self.start_button2.grid(row=1, column=1, pady=20)
        
        # Status messages
        self.status_label2 = ttk.Label(self.tab2, text="", foreground="green")
        self.status_label2.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

    def browse_source_dir1(self):
        directory = filedialog.askdirectory()
        if directory:
            self.source_dir_entry1.delete(0, tk.END)
            self.source_dir_entry1.insert(0, directory)

    def browse_dest_dir1(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dest_dir_entry1.delete(0, tk.END)
            self.dest_dir_entry1.insert(0, directory)

    def browse_dest_base_dir2(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dest_base_dir_entry2.delete(0, tk.END)
            self.dest_base_dir_entry2.insert(0, directory)

    def start_process1(self):
        # Get user inputs
        source_dir = self.source_dir_entry1.get()
        dest_dir = self.dest_dir_entry1.get()
        excel_path = r'Z:\Gene Synthesis\3.0 In-House Gene\3.5 Job Log\3.5 In-house Progress v3.XLSX' # Fixed excel path
        
        # Run the organizing script
        self.status_label1.config(text="Organizing files...")
        self.root.update_idletasks()
        self.organize_files(source_dir, dest_dir, excel_path)
        self.status_label1.config(text="Files organized successfully.")

    def start_process2(self):
        # Get user inputs
        dest_base_dir = self.dest_base_dir_entry2.get()
        
        # Run the distributing script
        self.status_label2.config(text="Distributing files...")
        self.root.update_idletasks()
        self.distribute_files(dest_base_dir)
        self.status_label2.config(text="Files distributed successfully.")

    def organize_files(self, source_dir, destination_dir, excel_path):
        if not os.path.exists(source_dir):
            print(f"Source directory does not exist: {source_dir}")
            return

        # Delete all .seq files in the source directory
        for filename in os.listdir(source_dir):
            file_path = os.path.join(source_dir, filename)
            if os.path.isfile(file_path) and filename.endswith('.seq'):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Error deleting file {file_path}: {e}")

        # Get the BBID mapping from the Excel file
        bbid_mapping = get_bbid_mapping(excel_path)

        # Loop through each file in the source directory to organize them
        for filename in os.listdir(source_dir):
            file_path = os.path.join(source_dir, filename)
            if os.path.isfile(file_path) and (filename.endswith('.ab1') or filename.endswith('.fasta')):
                job_id, plasmid_number = extract_info(filename)
                if job_id and plasmid_number:
                    # Get the corresponding BBID or set to empty string if unknown
                    bbid = bbid_mapping.get(job_id, '')
                    # Ensure no leading dots and extra dots in the folder name
                    folder_name = f'{job_id}.{plasmid_number}.{bbid}'.replace('..', '.')
                    folder_path = os.path.join(destination_dir, folder_name)
                    # Create the folder if it doesn't exist
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)
                    # Move the file to the folder
                    destination_file = os.path.join(folder_path, filename)
                    try:
                        shutil.move(file_path, destination_file)
                    except Exception as e:
                        print(f"Error moving file {file_path} to {destination_file}: {e}")

    def distribute_files(self, destination_base_dir):
        source_dir = r'Z:\Gene Synthesis\3.0 In-House Gene\3.6 QC\3.6.X Reference files'
        
        if not os.path.exists(source_dir):
            print(f"Source directory does not exist: {source_dir}")
            return

        # Create a dictionary to map base names to a list of full filenames
        file_map = {}
        for filename in os.listdir(source_dir):
            if filename.endswith('.txt'):
                base_name = filename.split('+')[0].split('.')[0]
                if base_name not in file_map:
                    file_map[base_name] = []
                file_map[base_name].append(filename)

        # Loop through each folder in the destination directory
        for folder in os.listdir(destination_base_dir):
            folder_base_name = folder.split('.')[0]  # Assume folder name starts with the base name
            if folder_base_name in file_map:
                destination_dir = os.path.join(destination_base_dir, folder)
                for source_filename in file_map[folder_base_name]:
                    source_file = os.path.join(source_dir, source_filename)
                    destination_file = os.path.join(destination_dir, source_filename)
                    try:
                        shutil.copy(source_file, destination_file)
                        print(f"Copied {source_filename} to {folder}")
                    except Exception as e:
                        print(f"Error copying file {source_file} to {destination_file}: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
