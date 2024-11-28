from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, 
                             QHBoxLayout, QLabel, QFileDialog, QTextEdit, QGridLayout)
from PyQt5.QtCore import Qt
import sys
import os
import pandas as pd
from datetime import datetime
import shutil

class CleanupGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('Data Cleanup Tool')
        self.setGeometry(100, 100, 600, 400)
        
        layout = QVBoxLayout()
        
        # Use QGridLayout for folder selection
        grid_layout = QGridLayout()
        
        # Uploaded folder selection
        self.uploaded_label = QLabel('Uploaded Folder:')
        self.uploaded_path = QLabel('Not selected')
        self.uploaded_path.setFixedWidth(300)  # Set fixed width
        self.uploaded_button = QPushButton('Select')
        self.uploaded_button.clicked.connect(lambda: self.select_folder('uploaded'))
        
        # Sorted sequencing folder selection
        self.sorted_label = QLabel('Sorted Sequencing Folder:')
        self.sorted_path = QLabel('Not selected')
        self.sorted_path.setFixedWidth(300)  # Set fixed width
        self.sorted_button = QPushButton('Select')
        self.sorted_button.clicked.connect(lambda: self.select_folder('sorted'))
        
        # Output folder selection
        self.output_label = QLabel('Output Folder:')
        self.output_path = QLabel('Not selected')
        self.output_path.setFixedWidth(300)  # Set fixed width
        self.output_button = QPushButton('Select')
        self.output_button.clicked.connect(lambda: self.select_folder('output'))
        
        # Add widgets to grid layout
        grid_layout.addWidget(self.uploaded_label, 0, 0)
        grid_layout.addWidget(self.uploaded_path, 0, 1)
        grid_layout.addWidget(self.uploaded_button, 0, 2)
        
        grid_layout.addWidget(self.sorted_label, 1, 0)
        grid_layout.addWidget(self.sorted_path, 1, 1)
        grid_layout.addWidget(self.sorted_button, 1, 2)
        
        grid_layout.addWidget(self.output_label, 2, 0)
        grid_layout.addWidget(self.output_path, 2, 1)
        grid_layout.addWidget(self.output_button, 2, 2)
        
        # Run button
        self.run_button = QPushButton('Run Cleanup')
        self.run_button.clicked.connect(self.run_cleanup)
        
        # Log area
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        
        layout.addLayout(grid_layout)
        layout.addWidget(self.run_button)
        layout.addWidget(self.log_area)
        
        self.setLayout(layout)
    
    def select_folder(self, folder_type):
        folder = QFileDialog.getExistingDirectory(self, f"Select {folder_type.capitalize()} Folder")
        if folder:
            if folder_type == 'uploaded':
                self.uploaded_path.setText(folder)
            elif folder_type == 'sorted':
                self.sorted_path.setText(folder)
            elif folder_type == 'output':
                self.output_path.setText(folder)
    
    def run_cleanup(self):
        uploaded_folder = self.uploaded_path.text()
        sorted_folder = self.sorted_path.text()
        output_folder = self.output_path.text()
        
        if 'Not selected' in (uploaded_folder, sorted_folder, output_folder):
            self.log_area.append("Please select all folders before running.")
            return
        
        self.log_area.append("Starting cleanup process...")
        
        excel_output_path = os.path.join(output_folder, f"{datetime.now().strftime('%Y-%m')} Project_Completion_Record.xlsx")
        cleanup_record_path = os.path.join(output_folder, f"{datetime.now().strftime('%Y-%m-%d')} Data Clean Up Record.xlsx")
        
        completed_projects = self.create_project_completion_excel(uploaded_folder, excel_output_path)
        self.cleanup_sorted_sequencing(sorted_folder, output_folder, completed_projects, cleanup_record_path)
        
        self.log_area.append("Cleanup process completed.")
    
    def create_project_completion_excel(self, folder, excel_path):
        self.log_area.append("Creating project completion Excel sheet...")
        data = []
        for file in os.listdir(folder):
            if file.endswith('.zip'):
                file_path = os.path.join(folder, file)
                date_created = datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
                data.append({'File Name': file, 'Date Created': date_created})
        df = pd.DataFrame(data)
        df['Work Number'] = df['File Name'].str[:9]
        df.to_excel(excel_path, index=False)
        self.log_area.append(f"Excel sheet created at: {excel_path}")
        return df
    
    def cleanup_sorted_sequencing(self, sorted_folder, output_folder, completed_projects, cleanup_record_path):
        self.log_area.append("Performing cleanup operation...")
        moved_folders = []
        
        for folder in os.listdir(sorted_folder):
            folder_path = os.path.join(sorted_folder, folder)
            if os.path.isdir(folder_path):
                work_number = folder[:9]  # Assuming the first 9 characters of the folder name are the work number
                if work_number in completed_projects['Work Number'].values:
                    destination_path = os.path.join(output_folder, folder)
                    shutil.move(folder_path, destination_path)
                    moved_folders.append({
                        'Folder Name': folder,
                        'Original Path': folder_path,
                        'New Path': destination_path,
                        'Date Moved': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    self.log_area.append(f"Moved folder: {folder} to {destination_path}")
        
        self.create_cleanup_record(cleanup_record_path, moved_folders)
    
    def create_cleanup_record(self, cleanup_record_path, moved_folders):
        self.log_area.append("Creating cleanup record...")
        df = pd.DataFrame(moved_folders)
        
        with pd.ExcelWriter(cleanup_record_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Moved Folders', index=False)
            
            # Add a summary sheet
            summary_data = {
                'Date': [datetime.now().strftime('%Y-%m-%d')],
                'Total Folders Moved': [len(moved_folders)]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        self.log_area.append(f"Cleanup record created at: {cleanup_record_path}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CleanupGUI()
    ex.show()
    sys.exit(app.exec_())