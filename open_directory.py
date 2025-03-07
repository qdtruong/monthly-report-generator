import os
import pandas as pd
from tkinter import Tk, filedialog

selected_folder = None  # Store the selected folder

def select_folder():
    global selected_folder
    if selected_folder:  # If already selected, return it
        return selected_folder

    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
#    root.deiconify()
#    root.update()
    selected_folder = filedialog.askdirectory(title="Select a Folder")
    root.destroy()

    if selected_folder:
        selected_folder = os.path.normpath(selected_folder)  # Ensure Windows format

    return selected_folder

def create_dataframe_from_folder(folder_path):
    if not folder_path:
        print("No folder selected.")
        return None
    
    files = [f for f in os.listdir(folder_path) if f.endswith(('.doc', '.docx'))]  # Filter for .doc and .docx files
    df_filenames = pd.DataFrame(files, columns=['Filename'])
    return df_filenames
