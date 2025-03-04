# -*- coding: utf-8 -*-
"""
Created on Mon Mar  3 19:30:38 2025

@author: quang.truong
"""
#import sys
#sys.path.append(r"C:\Users\quang.truong\OneDrive - HHS Office of the Secretary\Desktop\New folder")

import open_directory  # Imports python1.py
import monthly_report_extract_final  # Imports python2.py

# Run the function from python1.py to get the output dataframe
folder_path = open_directory.select_folder() # Replace `some_function` with the actual function that generates df
file_df = open_directory.create_dataframe_from_folder(folder_path) 

# Pass the output dataframe into python2.py's function
result = monthly_report_extract_final.combined_results(folder_path,file_df)  # Replace `some_other_function` with the function that accepts df_filenames


