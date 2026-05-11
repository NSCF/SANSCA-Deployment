# Digital Asset Metadata Sheet Generator (DAMSG)

A Python/Tkinter desktop tool for scanning SANSCA partner institution digital asset folders and generating DwC Simple Multimedia Extension–compliant metadata inventories for import into Living Atlas (LA) and AtoM archival systems.

## Features

- Scans TIFF, JPEG, RAW (NEF, CR2, CR3, ARW, DNG, ORF, RW2), PDF, and CSV files
- Loads collection mapping data live from a shared Google Sheet (no local file required)
- Supports three scan modes: Single Collection, All Collections for a selected institution, or All Institutions and Collections
- Labels each scan run with a storage type (Working Drive, Mirror Drive, Mirror RAID, Collection Copy, NAS Storage Repository)
- Generates SHA-256 checksums for every scanned file
- Extracts image capture dates via exiftool (with Pillow EXIF fallback)
- Outputs per-collection metadata CSVs for both LA and AtoM formats
- Accumulates a master digital asset inventory CSV (and optionally Excel) across scan runs
- Generates a preservation audit report comparing file presence and checksums across storage types
- Logs scan warnings (missing mappings, EXIF errors, checksum failures) to a separate CSV

## Expected Folder Structure

The root folder must follow this layout:

```
SANSCA_Root/
├── digital_vouchers/
│   └── <institutionCode>/
│       └── <collectionCode>/
│           ├── image_001.tif
│           └── ...
├── specimen_labels/
│   └── <institutionCode>/
│       └── <collectionCode>/
│           └── ...
├── registers/
│   └── <institutionCode>/
│       └── <collectionCode>/
│           └── ...
└── DAMSG_output/         ← created automatically
```

Categories (`digital_vouchers`, `specimen_labels`, `registers`) route to LA, AtoM, or both depending on configuration in the script.

## Requirements

Install Python 3 and the following packages:

```
pip install pandas openpyxl pillow
```

**exiftool** is strongly recommended for accurate image capture date extraction:

- Windows: download `exiftool(-k).exe` from [exiftool.org](https://exiftool.org), rename to `exiftool.exe`, place in `C:\Windows\`

If exiftool is not found on PATH the tool will warn you on startup and fall back to filesystem creation time.

## Usage

```
python digital_asset_metadata_sheet_generator_windows.py
```

### GUI Walkthrough

1. **Select SANSCA Root Folder** — choose the root folder containing your institution/collection subfolders.
2. **Load Mapping from Google Sheets** — fetches LA and AtoM collection mapping rows from the shared Google Sheet. This step is required; the tool will not proceed without a mapping loaded.
3. **Scan Type** — select the storage medium being scanned (Working Drive, Mirror Drive, Mirror RAID a.k.a Suzie, Collection Copy, NAS Storage Repository).
4. **Scan Mode** — choose Single Collection, All Collections for a selected institution, or All Institutions + Collections.
5. **Institution / Collection** — select the institution and collection to scan (enabled based on Scan Mode).
6. **File Filter** — limit scanning to All, TIFF only, RAW only, JPEG only, or PDF only.
7. **Output Choice** — select CSV only, Excel only, or Both.
8. **Clear previous metadata files** *(testing only)* — deletes existing per-collection metadata CSVs before the scan.
9. **Clear master inventory files** *(testing only)* — deletes existing master inventory and audit CSVs before the scan.
10. **Start Processing** — closes the GUI and begins the scan.

## Output Files

All output is written to a `DAMSG_output/` folder inside the selected root folder.

```
DAMSG_output/
├── digital_asset_inventory_la_<timestamp>.csv        ← master LA inventory (cumulative)
├── digital_asset_inventory_la_<timestamp>.xlsx       ← optional Excel version
├── digital_asset_inventory_atom_<timestamp>.csv      ← master AtoM import CSV (cumulative)
├── preservation_audit_la_<timestamp>_summary.csv     ← cross-storage comparison summary
├── preservation_audit_la_<timestamp>_missing.csv     ← files missing on target storage
├── preservation_audit_la_<timestamp>_mismatch.csv    ← checksum mismatches between storages
├── preservation_audit_la_<timestamp>_duplicates.csv  ← duplicate files (same checksum)
├── preservation_audit_atom_<timestamp>_*.csv         ← same audit files for AtoM items
├── scan_warnings_<timestamp>.csv                     ← EXIF errors, missing mappings, etc.
└── metadata/
    └── <institutionCode>/
        └── <collectionCode>/
            ├── <collectionCode>_<category>_metadata_la_<timestamp>.csv
            └── <collectionCode>_<category>_metadata_atom_<timestamp>.csv
```

Master inventory files are cumulative: each run appends new rows (deduplicated by `documentId` + `scanType`), so you can scan different storage types over time and build up a complete cross-storage picture.

## Preservation Audit

The audit compares files across the following logical storage pairs:

| Source | Target |
|--------|--------|
| Working Drive | Mirror Drive |
| Mirror Drive | Mirror RAID a.k.a Suzie |
| Working Drive | Collection Copy |
| Mirror RAID a.k.a Suzie | NAS Storage Repository |

For each pair the audit reports: total source files, matching files, files missing on target, and SHA-256 checksum mismatches.

## Creating a Standalone Executable

1. Install PyInstaller:

```
pip install pyinstaller
```

2. From the directory containing the script, run:

```
pyinstaller --onefile --windowed digital_asset_metadata_sheet_generator_windows.py
```

3. The executable will be located at:

```
dist/digital_asset_metadata_sheet_generator_windows.exe
```

## License

MIT License © 2025
