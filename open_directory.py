import os
import pandas as pd
from tkinter import Tk, filedialog

def select_folder():
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring the window to the front
    root.lift()  # Lift the window to the top
    root.focus_force()  # Force focus on the window
    folder_selected = filedialog.askdirectory(title="Select a Folder")
    root.attributes('-topmost', False)  # Reset window attribute
    if folder_selected:
        folder_selected = os.path.normpath(folder_selected)  # Convert to Windows-compatible format
    return folder_selected

def create_dataframe_from_folder(folder_path):
    if not folder_path:
        print("No folder selected.")
        return None
    
    files = [f for f in os.listdir(folder_path) if f.endswith(('.doc', '.docx'))]  # Filter for .doc and .docx files
    df_filenames = pd.DataFrame(files, columns=['Filename'])
    return df_filenames

if __name__ == "__main__":
    folder = select_folder()
    df_filenames = create_dataframe_from_folder(folder)
    
    if df_filenames is not None:
        print(df_filenames)  # Display the DataFrame
        # Optionally save to a CSV file
        # df_filenames.to_csv("file_list.csv", index=False)
