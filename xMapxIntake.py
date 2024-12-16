import os
import shutil
from datetime import datetime
import pandas as pd
import pyodbc
import tkinter as tk
from tkinter import filedialog

# General Utilities

def ensure_folder_exists(folder_path):
    """Create folder if it doesn't exist."""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


# Folder Structure Creation

def create_folder_structure(base_dir):
    """Create a folder structure for today's date with Data and Documents subfolders."""
    today = datetime.today().strftime('%Y%m%d')
    root_folder = os.path.join(base_dir, today)

    data_folder = os.path.join(root_folder, 'Data')
    documents_folder = os.path.join(root_folder, 'Documents')

    for folder in [root_folder, data_folder, documents_folder]:
        ensure_folder_exists(folder)

    return root_folder, data_folder, documents_folder


# XJMTR Processing for Judgment PDFs

def process_xjmtr(src_dir, xjmtr_file_path):
    """Process PDFs with 'judgment' in the filename to generate xjmtr.txt."""
    with open(xjmtr_file_path, 'w') as xjmtr_file:
        xjmtr_file.write("FILENO\tLLCODE\n")  # Write header

        for filename in os.listdir(src_dir):
            if filename.lower().endswith('.pdf') and "judgment" in filename.lower():
                fileno = filename.split('_')[0]
                if fileno.isdigit():
                    xjmtr_file.write(f"{fileno}\tXJMTR\n")
                    print(f"Added to xjmtr.txt: {fileno}\tXJMTR")


# Database Connection and Query

def connect_to_sql():
    """Connect to the SQL database."""
    connection_string = "DRIVER={SQL Server};SERVER=your_server_name;DATABASE=enterdatabasehere;Trusted_Connection=yes;"
    return pyodbc.connect(connection_string)

def fetch_defendants(person_served):
    """Fetch defendants from the database based on the person served."""
    query = """
    SELECT FileNo, LTRIM(RTRIM(DEFENDANT_1)), LTRIM(RTRIM(DEFENDANT_2)), LTRIM(RTRIM(DEFENDANT_3))
    FROM dbo.MASTER
    WHERE DEFENDANT_1 IS NOT NULL OR DEFENDANT_2 IS NOT NULL OR DEFENDANT_3 IS NOT NULL
    """
    with connect_to_sql() as conn:
        cursor = conn.cursor()
        cursor.execute(query)
        for row in cursor.fetchall():
            file_no, d1, d2, d3 = row
            if person_served in [d1, d2, d3]:
                return file_no, d1, d2, d3
    return None, None, None, None


# Excel File Processing

def process_excel_file(input_excel_path):
    """Process Excel file and extract data."""
    df = pd.read_excel(input_excel_path, engine='openpyxl')
    df.columns = df.columns.str.strip()

    if 'Document Category' not in df.columns:
        print(f"'Document Category' column missing in {input_excel_path}")
        return [], []

    filtered_df = df[df['Document Category'] == 'Summons and Complaint']
    output_data_1 = []
    output_data_2 = []
    processed_pairs = set()

    for _, row in filtered_df.iterrows():
        file_no = row['FileNo']
        person_served = row['Person Served']
        service_type = row['Service Type']
        date_of_service = row['Date of Service']
        address = row['Service Street Address']
        notes = row.get('Note', '')

        if (file_no, person_served) in processed_pairs:
            continue

        matched_file_no, d1, d2, d3 = fetch_defendants(person_served)
        if matched_file_no:
            llcode = determine_llcode(service_type, person_served, d1, d2, d3)
            if llcode:
                output_data_1.append([file_no, llcode, person_served, date_of_service, address, service_type])
                cleaned_notes = notes.replace('\n', ' ').replace('\r', ' ')
                output_data_2.append([109, 'D', file_no, 'XNSRV', 'MWC', cleaned_notes, '#'])
                processed_pairs.add((file_no, person_served))

    return output_data_1, output_data_2

def determine_llcode(service_type, person_served, d1, d2, d3):
    """Determine LLCode based on service type and defendants."""
    if service_type == "N":
        return f"XSNG{[d1, d2, d3].index(person_served) + 1}" if person_served in [d1, d2, d3] else None
    elif service_type == "S":
        return f"XSUBS{[d1, d2, d3].index(person_served) + 1}" if person_served in [d1, d2, d3] else None
    elif service_type == "P":
        return f"XPS{[d1, d2, d3].index(person_served) + 1}" if person_served in [d1, d2, d3] else None
    return None


# File Movement and Grouping

def move_files(src_dir, data_folder, documents_folder, xjmtr_file_path):
    """Organize files into folders and process Excel and PDF files."""
    process_xjmtr(src_dir, xjmtr_file_path)

    for filename in os.listdir(src_dir):
        file_path = os.path.join(src_dir, filename)
        if os.path.isfile(file_path):
            name, ext = os.path.splitext(filename)
            if ext.lower() in ['.xlsx', '.xls', '.csv']:
                folder_name = name.split('_')[0]
                subfolder = os.path.join(data_folder, folder_name)
                ensure_folder_exists(subfolder)
                shutil.move(file_path, os.path.join(subfolder, filename))
            elif ext.lower() == '.pdf':
                folder_name = name.split('_')[1] if len(name.split('_')) > 1 else 'Unknown'
                subfolder = os.path.join(documents_folder, folder_name)
                ensure_folder_exists(subfolder)
                new_filename = filename.replace('_', ' ')
                shutil.move(file_path, os.path.join(subfolder, new_filename))


# Main Processing

def select_and_process_folder():
    """Select folder and process all files."""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Folder with Files")

    if not folder_path:
        print("No folder selected.")
        return

    root_folder, data_folder, documents_folder = create_folder_structure(folder_path)
    xjmtr_file_path = os.path.join(root_folder, "xjmtr.txt")

    move_files(folder_path, data_folder, documents_folder, xjmtr_file_path)
    print(f"Processing complete. Files organized in: {root_folder}")


# Main Script Entry
if __name__ == "__main__":
    select_and_process_folder()
