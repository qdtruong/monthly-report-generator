# -*- coding: utf-8 -*-
"""
Created on Wed Mar  5 10:25:39 2025

@author: quang.truong
"""


#result2 = result[result.Detail != ""]
#result2 = result2[result2.Staff == "Sadia Rahman"]



from docx import Document
from docx.shared import Pt

# Load the existing Word document
doc_path = r"C:\Users\quang.truong\OneDrive - HHS Office of the Secretary\Desktop\Report Generator\Template.docx"
doc = Document(doc_path)

# List of tasks from the DataFrame
tasks = [
    "Check emails", "Attend team meeting", "Prepare monthly report", "Client call",
    "Update project plan", "Submit expense report", "Write presentation slides", "Conduct market research",
    "Schedule meetings", "Review documents", "Provide team feedback", "Fix spreadsheet issues",
    "Organize files", "", "Follow up with clients", "Train new hire",
    "Analyze sales data", "Create social media posts", "Update CRM", "Plan team event"
]

# Find the second column in the document (assuming a table exists)
tables = doc.tables
if tables:
    table = tables[0]  # Assuming we are modifying the first table
    cell = table.rows[1].cells[2]  # second row, middle column (I think there's hidden columns)
#    cell = table.rows[1].cells[4]  # second row, right column (I think there's hidden columns)
    
    # Create a bulleted list as a single text block
    bullet_list = "\n".join([f"• {task}" for task in tasks])
    
    p = cell.paragraphs[0]
    p.text = bullet_list
    run = p.runs[0]
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)

# Save the modified document
output_path = r"C:\Users\quang.truong\OneDrive - HHS Office of the Secretary\Desktop\Report Generator\Template2.docx"
doc.save(output_path)

#output_path
