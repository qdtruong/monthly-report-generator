# -*- coding: utf-8 -*-
"""
Created on Mon Mar  3 09:29:04 2025

@author: quang.truong
"""
#pip install pywin32
import win32com.client
import pandas as pd
from datetime import datetime

# Define file path
doc_path = r"C:\Users\quang.truong\OneDrive - HHS Office of the Secretary\Desktop\New folder\25-02-monthly-report-dsl.docx"

# Open Word application
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # Keep Word application hidden
doc = word.Documents.Open(doc_path)

# Extract table data
data = []
for table in doc.Tables:
    num_rows = table.Rows.Count

    for row_idx in range(2, num_rows + 1):  # Skip header row, start from second row
        staff = table.Cell(row_idx, 2).Range.Text.strip()  # Staff name
        activity_text = table.Cell(row_idx, 3).Range.Text.strip()  # Activity/Success details
        pending_text = table.Cell(row_idx, 4).Range.Text.strip()  # Pending Actions details

        # Normalize text (remove unwanted Word artifacts)
        staff = staff.replace("\r", "").replace("\n", "").strip()
        activity_text = activity_text.replace("\r", "\n").strip()  # Preserve line breaks
        pending_text = pending_text.replace("\r", "\n").strip()

        # Function to clean and extract bullet points
        def extract_bullets(text):
            bullets = [line.strip("-•").strip() for line in text.split("\n") if line.strip()]
            return bullets

        # Process Activity/Success
        for activity in extract_bullets(activity_text):
            data.append([datetime.today().strftime('%Y-%m-%d'), staff, activity, "Activity/Success"])

        # Process Pending Actions
        for pending in extract_bullets(pending_text):
            data.append([datetime.today().strftime('%Y-%m-%d'), staff, pending, "Pending Actions"])

# Close Word document
doc.Close()
word.Quit()

# Convert to DataFrame
df = pd.DataFrame(data, columns=["Date", "Staff", "Detail", "Category"])
