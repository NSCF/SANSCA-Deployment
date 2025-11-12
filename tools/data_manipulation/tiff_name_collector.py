import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox

# Hide the root window
TK().withdraw()

# Ask user to select the main folder
main_folder = filedialog.askdirectory(title="Hey Jana, please select the Main Folder to Search")
if not main_folder:
        messagebox.showinfo("Cancelled", "No folder selected. Exiting")
        exit()

data = [] 

# Walk through folders recursively
for root, dirs, files in os.walk(main_folder):
        for files in files:
                if file.lower().endswith(('.tiff')):
                        folder_name = os.path.basename(root)
                        data.append({
                                'Folder Name': folder_name,
                                'File Name': file,
                                'Full Path': os.path.join(root,file)
                        })

if not data:
        messagebox.showinfo("Done", "No TIFF files found in the selected folder.")
        exit() 

# Save to excel
df = pd.DataFrame(data)
output_path = os.path,join(main_folder, "tiff_files_list.xlsx",)
df.to_excel(output_path, index=False)

messagebox.showinfo("Success", "Excel file created:\n\n{output_path}")
