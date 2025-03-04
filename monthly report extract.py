# -*- coding: utf-8 -*-
"""
Created on Mon Mar  3 15:02:25 2025

@author: quang.truong
"""

import win32com.client
import pandas as pd
from datetime import datetime
import os

def extract_data_from_word(file_path):
    """Extracts data from a Word document table and returns a DataFrame."""
    
    # Open Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Keep Word application hidden
    doc = word.Documents.Open(file_path)

    # Extract table data
    data = []
    for table in doc.Tables:
        num_rows = table.Rows.Count
        num_cols = table.Columns.Count  # Check number of columns

        for row_idx in range(2, num_rows + 1):  # Skip header row, start from second row
            try:
                if num_cols == 4:
                    # If there are 4 columns, ignore the first column
                    staff = table.Cell(row_idx, 2).Range.Text.strip()  # Staff name
                    activity_text = table.Cell(row_idx, 3).Range.Text.strip()  # Activity/Success details
                    pending_text = table.Cell(row_idx, 4).Range.Text.strip()  # Pending Actions details
                elif num_cols == 3:
                    # Standard format: 3 columns
                    staff = table.Cell(row_idx, 1).Range.Text.strip()  # Staff name
                    activity_text = table.Cell(row_idx, 2).Range.Text.strip()  # Activity/Success details
                    pending_text = table.Cell(row_idx, 3).Range.Text.strip()  # Pending Actions details
                else:
                    continue  # Skip tables that don't match expected formats
            except:
                continue  # Skip row if there's an error

            # Normalize text (remove unwanted Word artifacts)
            staff = staff.replace("\r", "").replace("\n", "").replace("", "").strip()
            activity_text = activity_text.replace("\r", "\n").strip()  # Preserve line breaks
            pending_text = pending_text.replace("\r", "\n").strip()

            # Function to clean and extract bullet points
            def extract_bullets(text):
                bullets = [line.strip("-•").strip() for line in text.split("\n") if line.strip()]
                return bullets

            # Process Activity/Success
            for activity in extract_bullets(activity_text):
                data.append([datetime.today().strftime('%Y-%m-%d'), staff, activity, "Activity/Success", os.path.basename(file_path)])

            # Process Pending Actions
            for pending in extract_bullets(pending_text):
                data.append([datetime.today().strftime('%Y-%m-%d'), staff, pending, "Pending Actions", os.path.basename(file_path)])

    # Close Word document
    doc.Close()
    word.Quit()

    # Convert to DataFrame
    return pd.DataFrame(data, columns=["Date", "Staff", "Detail", "Category", "Source File"])

# Example DataFrame with file names
folder_path = r"C:\Users\quang.truong\OneDrive - HHS Office of the Secretary\Desktop\New folder"
file_df = pd.DataFrame({
    "Filename": ["Monthly Report_BritnieBarrett.doc", "2-28-25 Monthly Report.doc", "25-02-monthly-report-dsl.docx"]
})

def combined_results(folder_path,file_df):
# Process all files and combine results
    all_data = pd.DataFrame()
    
    for file_name in file_df["Filename"]:
        file_path = os.path.join(folder_path, file_name)
        
        if os.path.exists(file_path):  # Ensure the file exists before processing
            df = extract_data_from_word(file_path)
            all_data = pd.concat([all_data, df], ignore_index=True)
        else:
            print(f"Skipping: {file_name} (File not found)")

    return all_data

# Save final combined output
#output_path = "/mnt/data/combined_extracted_report.xlsx"
#all_data.to_excel(output_path, index=False)

#print(f"All data successfully extracted and saved to {output_path}")
final_output = combined_results(folder_path,file_df)
