## TIFF Folder Scanner
This Python tool scans a selected folder (and all its subfolders) for TIFF image files (.tif / .tiff) and generates an Excel spreadsheet listing:
* Institution name
* Folder name
* File name
* Full file path
* Relative file path

It uses a simple Tkinter GUI for folder selection and pop-up messages.

## Features
* Recursively scans all subfolders for TIFF files
* Automatically includes folder, filename, and path details
* Saves the results to an Excel file named tiff_files_list.xlsx in the selected folder
* User-friendly pop-up dialogs for instructions, errors, and completion messages

## Requirements
You’ll need Python 3 installed, along with the following dependencies:
```
pip install pandas openpyxl
```
**Note:** openpyxl is required for writing .xlsx files.

## Usage
1. Download or clone this repository.
2. Run the script in your terminal or IDE:
```
python3 tiff_folder_scanner_macos.py
```
3. When prompted, select the main institution folder containing all your collection subfolders.
4. The tool will search recursively for .tif and .tiff files.
5. Once complete, an Excel file named **tiff_files_list.xlsx** will be saved in the same main folder.

## Folder Structure Example
```
Institution_Folder/
├── Collection_A/
│   ├── image_001.tif
│   ├── image_002.tiff
├── Collection_B/
│   ├── sample_01.tif
│   ├── sample_02.tif
```
Output Excel Example:
| Institution Name | Folder Name |	File Name |	Full Path |	Relative Path |
|------------------|-------------|------------|-----------|---------------|
| Institution_Folder |	Collection_A |	image_001.tif |	/path/to/Collection_A/image_001.tif |	Collection_A/image_001.tif
| Institution_Folder |	Collection_B |	sample_02.tif |	/path/to/Collection_B/sample_02.tif |	Collection_B/sample_02.tif

## Screenshots (Update)

1. Folder selection dialog (Add)
2. Success message (Add)
3. Example Excel output (Add)

**Note:** Save your screenshots in a folder named screenshots/ inside your repository. (Create folder)

## Notes
* The Excel output can be opened in Excel, LibreOffice Calc, or Google Sheets.
* If you see an error related to permissions or open files, make sure the Excel file isn’t already open.
* Works on Windows, macOS, and Linux (with a GUI environment).

## Creating a Standalone Executable

This utility can be packaged as a standalone executable, if required.

1. Install PyInstaller:
```
pip install pyinstaller
```
2. From the directory containing tiff_folder_scanner_windows.py, run:
```
pyinstaller --onefile --windowed tiff_folder_scanner_windows.py
```
3. After the build finishes, your executable will be located at:
```
dist/tiff_folder_scanner_macos
```

## License
MIT License © 2025

