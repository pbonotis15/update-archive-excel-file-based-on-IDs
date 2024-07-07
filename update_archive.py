# Author: Panos Bonotis -> https://www.linkedin.com/in/panagiotis-bonotis-351a7996/
# Date: Jul-2024
# Description: This script is designed to merge new SR IDs from a daily received excel file 
# into an existing "archive" file. It checks multiple specified sheets for new SR IDs 
# and appends only the new entries to the existing file, ensuring that any notes or existing data 
# in the original file are preserved.

import pandas as pd
from tkinter import filedialog, Tk

def get_file_path(title):
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def append_new_sr_ids(existing_file_path, new_file_path, sheets):
    # Read the existing and new final_results files for each sheet
    with pd.ExcelWriter(existing_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for sheet in sheets:
            existing_df = pd.read_excel(existing_file_path, sheet_name=sheet)
            new_df = pd.read_excel(new_file_path, sheet_name=sheet)
            
            # Identify new SR IDs
            new_sr_ids = new_df[~new_df['SR ID'].isin(existing_df['SR ID'])]
            
            # Append new SR IDs to the existing dataframe
            updated_df = pd.concat([existing_df, new_sr_ids], ignore_index=True)
            
            # Save the updated dataframe back to the existing file
            updated_df.to_excel(writer, sheet_name=sheet, index=False)

if __name__ == "__main__":
    existing_file_path = get_file_path("Select the existing final_results file")
    new_file_path = get_file_path("Select the new final_results file")

    # Specify the sheets you want to check
    sheets_to_check = ["Aggregated Data", "Summary of Actions", "Last Drop"]

    if existing_file_path and new_file_path:
        append_new_sr_ids(existing_file_path, new_file_path, sheets_to_check)
        print("New SR IDs have been appended successfully to all specified sheets.")
    else:
        print("File selection was cancelled.")