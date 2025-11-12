#!/usr/bin/env python3
import os
import sys
import pandas as pd
from tkinter import Tk, filedialog, messagebox

# Create the root window
root = Tk()
root.withdraw()  # Hide the main window

# Force dialogs to appear on top
root.attributes("-topmost", True)
root.update() # Make sure the attribute takes effect

# Intro message
messagebox.showinfo(
    "Welcome",
    "Welcome!\n\n"
    "This tool will scan the folder you select for all TIFF files (.tif, .tiff) "
    "and create an Excel list that includes folder names, file names, full paths, and relative paths.\n\n"
    "IMPORTANT:\n"
    "- Each file should be stored in a folder named according to its collection.\n"
    "- All collection folders must be placed inside the main institution folder you select."
)

# Ask user to select the main folder
main_folder = filedialog.askdirectory(title="Please select the Main Institution Folder to Search")
if not main_folder:
    messagebox.showinfo("Cancelled", "No folder selected. Exiting.")
    root.destroy()
    sys.exit()

data = []

# Get institution name from main folder
institution_name = os.path.basename(main_folder)

# Walk through folders recursively
for r, dirs, files in os.walk(main_folder):
    for file in files:
        if file.lower().endswith(('.tiff', '.tif')):
            folder_name = os.path.basename(r)
            full_path = os.path.join(r, file)
            rel_path = os.path.relpath(full_path, start=main_folder)  # Relative path
            data.append({
                'Institution Name': institution_name,
                'Folder Name': folder_name,
                'File Name': file,
                'Full Path': full_path,
                'Relative Path': rel_path
            })

# Check if any TIFF files were found
if not data:
    messagebox.showinfo("Done", "No TIFF files found in the selected folder.")
    sys.exit()

# Save to Excel with error handling
df = pd.DataFrame(data)
output_path = os.path.join(main_folder, "tiff_files_list.xlsx")
try:
    df.to_excel(output_path, index=False)
    messagebox.showinfo("Success", f"Excel file created successfully!\n\n{output_path}")
except Exception as e:
    messagebox.showerror("Error", f"Could not save Excel file:\n{e}")

# Clean up Tkinter
root.destroy()