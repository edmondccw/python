import os
import shutil
import pandas as pd
from datetime import date
from zipfile import ZipFile
import zipfile
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QLineEdit, QLabel, QTextEdit, QFileDialog, QTabWidget
from PyQt5.QtCore import QThread, pyqtSignal

class RenameWorkerThread(QThread):
    update_signal = pyqtSignal(str)
    
    def __init__(self, main_folder, excel_path):
        QThread.__init__(self)
        self.main_folder = main_folder
        self.excel_path = excel_path

    def run(self):
        self.process_folders()

    def process_folders(self):
        # Function to extract work number from folder name
        def extract_work_number(folder_name):
            return folder_name[:9]

        # Function to get today's date in the format YYYY-MM-DD
        def get_today_date():
            return date.today().strftime("%Y-%m-%d")

        # Read the reference Excel file
        try:
            ref_df = pd.read_excel(self.excel_path)
        except FileNotFoundError:
            self.update_signal.emit(f"Error: Excel file not found at {self.excel_path}")
            return

        # Print column names for debugging
        self.update_signal.emit(f"Columns in the Excel file: {ref_df.columns.tolist()}")

        # Strip whitespace from column names
        ref_df.columns = ref_df.columns.str.strip()

        # Check if required columns exist
        required_columns = ['Work Number', 'BBID', 'Vector']
        missing_columns = [col for col in required_columns if col not in ref_df.columns]
        if missing_columns:
            self.update_signal.emit(f"Error: The following required columns are missing from the Excel file: {', '.join(missing_columns)}")
            return

        # Convert the Work Number column to string type
        ref_df['Work Number'] = ref_df['Work Number'].astype(str)

        # Check if the main folder exists
        if not os.path.exists(self.main_folder):
            self.update_signal.emit(f"Error: The main folder does not exist at {self.main_folder}")
            return

        # Iterate through all subfolders in the main folder
        for folder_name in os.listdir(self.main_folder):
            folder_path = os.path.join(self.main_folder, folder_name)
           
            # Check if it's a directory
            if os.path.isdir(folder_path):
                # Extract work number from the folder name
                work_number = extract_work_number(folder_name)
               
                # Find matching row in the reference dataframe
                matching_row = ref_df[ref_df['Work Number'] == work_number]
               
                if not matching_row.empty:
                    bbid = matching_row['BBID'].values[0]
                    vector = matching_row['Vector'].values[0]
                   
                    # Create new folder name
                    new_folder_name = f"{work_number}.{bbid} in {vector} {get_today_date()}"
                   
                    # Full path for the new folder name
                    new_folder_path = os.path.join(self.main_folder, new_folder_name)
                   
                    # Rename the folder
                    try:
                        os.rename(folder_path, new_folder_path)
                        self.update_signal.emit(f"Renamed: {folder_name} -> {new_folder_name}")
                    except PermissionError:
                        self.update_signal.emit(f"Error: Permission denied when trying to rename {folder_name}")
                    except FileExistsError:
                        self.update_signal.emit(f"Error: A folder with the name {new_folder_name} already exists")
                else:
                    self.update_signal.emit(f"No matching work number found for: {work_number}")

        self.update_signal.emit("Folder renaming completed.")

class ZipWorkerThread(QThread):
    update_signal = pyqtSignal(str)
    
    def __init__(self, folder_directory):
        QThread.__init__(self)
        self.folder_directory = folder_directory

    def run(self):
        self.zip_folders()

    def zip_folders(self):
        def zip_folder(folder_path, zip_name):
            with ZipFile(zip_name, 'w') as zipf:
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, start=folder_path)
                        zipf.write(file_path, arcname)

        for folder_name in os.listdir(self.folder_directory):
            folder_path = os.path.join(self.folder_directory, folder_name)
           
            if os.path.isdir(folder_path):
                zip_name = os.path.join(self.folder_directory, f"{folder_name}.zip")
               
                try:
                    zip_folder(folder_path, zip_name)
                    self.update_signal.emit(f"Zipped '{folder_name}' to '{zip_name}'")
                except Exception as e:
                    self.update_signal.emit(f"Error zipping '{folder_name}': {e}")
                    continue
               
                try:
                    shutil.rmtree(folder_path)
                    self.update_signal.emit(f"Removed original folder '{folder_name}'")
                except Exception as e:
                    self.update_signal.emit(f"Error removing folder '{folder_name}': {e}")

        self.update_signal.emit("Zipping complete.")

class UnzipWorkerThread(QThread):
    update_signal = pyqtSignal(str)
    
    def __init__(self, directory):
        QThread.__init__(self)
        self.directory = directory

    def run(self):
        self.unzip_files()

    def unzip_files(self):
        for filename in os.listdir(self.directory):
            if filename.endswith('.zip'):
                zip_file = os.path.join(self.directory, filename)
                folder_name = os.path.splitext(os.path.basename(zip_file))[0]
                extract_folder = os.path.join(self.directory, folder_name)
                os.makedirs(extract_folder, exist_ok=True)
                
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    zip_ref.extractall(extract_folder)
                
                contents = os.listdir(extract_folder)
                if len(contents) == 1 and os.path.isdir(os.path.join(extract_folder, contents[0])):
                    inner_folder = os.path.join(extract_folder, contents[0])
                    for item in os.listdir(inner_folder):
                        shutil.move(os.path.join(inner_folder, item), extract_folder)
                    os.rmdir(inner_folder)
                
                os.remove(zip_file)
                self.update_signal.emit(f"Extracted '{filename}' to '{extract_folder}' and deleted '{filename}'.")

        self.update_signal.emit("Unzipping complete.")

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Folder Operations'
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle(self.title)
        
        layout = QVBoxLayout()
        
        # Create tabs
        self.tabs = QTabWidget()
        self.rename_tab = QWidget()
        self.zip_tab = QWidget()
        self.unzip_tab = QWidget()
        
        self.tabs.addTab(self.rename_tab, "Rename")
        self.tabs.addTab(self.zip_tab, "Zip")
        self.tabs.addTab(self.unzip_tab, "Unzip")
        
        layout.addWidget(self.tabs)
        
        # Rename tab
        rename_layout = QVBoxLayout()
        self.main_folder_input = QLineEdit()
        self.excel_file_input = QLineEdit()
        
        rename_layout.addWidget(QLabel('Main Folder:'))
        main_folder_layout = QHBoxLayout()
        main_folder_layout.addWidget(self.main_folder_input)
        main_folder_layout.addWidget(QPushButton('Browse', clicked=lambda: self.browse_folder(self.main_folder_input)))
        rename_layout.addLayout(main_folder_layout)
        
        rename_layout.addWidget(QLabel('Reference Excel File:'))
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_file_input)
        excel_layout.addWidget(QPushButton('Browse', clicked=lambda: self.browse_file(self.excel_file_input)))
        rename_layout.addLayout(excel_layout)
        
        self.rename_button = QPushButton('Rename Folders', clicked=self.run_rename_script)
        rename_layout.addWidget(self.rename_button)
        
        self.rename_tab.setLayout(rename_layout)
        
        # Zip tab
        zip_layout = QVBoxLayout()
        self.zip_folder_input = QLineEdit()
        
        zip_layout.addWidget(QLabel('Folder to Zip:'))
        zip_folder_layout = QHBoxLayout()
        zip_folder_layout.addWidget(self.zip_folder_input)
        zip_folder_layout.addWidget(QPushButton('Browse', clicked=lambda: self.browse_folder(self.zip_folder_input)))
        zip_layout.addLayout(zip_folder_layout)
        
        self.zip_button = QPushButton('Zip Folders', clicked=self.run_zip_script)
        zip_layout.addWidget(self.zip_button)
        
        self.zip_tab.setLayout(zip_layout)
        
        # Unzip tab
        unzip_layout = QVBoxLayout()
        self.unzip_folder_input = QLineEdit()
        
        unzip_layout.addWidget(QLabel('Folder to Unzip:'))
        unzip_folder_layout = QHBoxLayout()
        unzip_folder_layout.addWidget(self.unzip_folder_input)
        unzip_folder_layout.addWidget(QPushButton('Browse', clicked=lambda: self.browse_folder(self.unzip_folder_input)))
        unzip_layout.addLayout(unzip_folder_layout)
        
        self.unzip_button = QPushButton('Unzip Folders', clicked=self.run_unzip_script)
        unzip_layout.addWidget(self.unzip_button)
        
        self.unzip_tab.setLayout(unzip_layout)
        
        # Log output
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)
        
        self.setLayout(layout)
        self.show()
        
    def browse_folder(self, input_field):
        folder = QFileDialog.getExistingDirectory(self, "Select Directory")
        if folder:
            input_field.setText(folder)
            
    def browse_file(self, input_field):
        file, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Excel Files (*.xlsx *.xls)")
        if file:
            input_field.setText(file)
            
    def run_rename_script(self):
        main_folder = self.main_folder_input.text()
        excel_file = self.excel_file_input.text()
        
        if not all([main_folder, excel_file]):
            self.log_output.append("Please fill in all fields for renaming.")
            return
        
        self.rename_worker = RenameWorkerThread(main_folder, excel_file)
        self.rename_worker.update_signal.connect(self.update_log)
        self.rename_worker.start()
        self.rename_button.setEnabled(False)
        
    def run_zip_script(self):
        folder_directory = self.zip_folder_input.text()
        
        if not folder_directory:
            self.log_output.append("Please select a folder to zip.")
            return
        
        self.zip_worker = ZipWorkerThread(folder_directory)
        self.zip_worker.update_signal.connect(self.update_log)
        self.zip_worker.start()
        self.zip_button.setEnabled(False)
        
    def run_unzip_script(self):
        directory = self.unzip_folder_input.text()
        
        if not directory:
            self.log_output.append("Please select a folder to unzip.")
            return
        
        self.unzip_worker = UnzipWorkerThread(directory)
        self.unzip_worker.update_signal.connect(self.update_log)
        self.unzip_worker.start()
        self.unzip_button.setEnabled(False)
        
    def update_log(self, message):
        self.log_output.append(message)
        if message.endswith("completed."):
            self.rename_button.setEnabled(True)
            self.zip_button.setEnabled(True)
            self.unzip_button.setEnabled(True)

if __name__ == '__main__':
    app = QApplication([])
    ex = App()
    app.exec_()