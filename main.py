# -*- coding: utf-8 -*-
"""
Created on Mon Mar  3 19:30:38 2025

@author: quang.truong
"""

#this only needs to be run if running code from Spyder interpreter (instead of running with Run File/F5)
#import sys
#sys.path.append(r"C:\Users\quang.truong\OneDrive - HHS Office of the Secretary\Desktop\Report Generator")

import open_directory  # Imports module 1
import monthly_report_extract_final  # Imports module 2
import report_output # Imports module 3

# Run the function from module 1 to get the output dataframe
folder_path = open_directory.select_folder() # Replace `some_function` with the actual function that generates df
file_df = open_directory.create_dataframe_from_folder(folder_path) 

# Pass the output dataframe into module 2's function
result = monthly_report_extract_final.combined_results(folder_path,file_df)  # Replace `some_other_function` with the function that accepts df_filenames

# Output report with module 3
report_output.report_output(result, folder_path)

