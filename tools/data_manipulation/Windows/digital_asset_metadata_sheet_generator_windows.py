#!/usr/bin/env python3
import os
import subprocess
import pandas as pd
from datetime import datetime
from tkinter import (
    Tk, Canvas, Label, Button, StringVar, BooleanVar,
    OptionMenu, Checkbutton, filedialog, messagebox, DISABLED, NORMAL
)
from PIL import Image, ExifTags, ImageTk
import platform
import sys
import pathlib
import hashlib

# ==================================================
# Google Sheets configuration
# ==================================================
SHEET_ID = "1AVqVoy8Hvk3GpJ0mCMXHjbZwDvMOGxa50CdH0Jh6bOU"
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/{id}/gviz/tq?tqx=out:csv&sheet={sheet}".format

# ==================================================
# Category → output target routing
# ==================================================
CATEGORY_TARGETS = {
    "digital_vouchers": ["LA"],
    "specimen_labels":  ["LA", "AtoM"],
    "registers":        ["AtoM"],
}

# ==================================================
# AtoM CSV column order (matches master_atom sheet)
# ==================================================
ATOM_COLUMNS = [
    "legacyId", "parentId", "qubitParentSlug", "accessionNumber",
    "identifier", "title", "levelOfDescription", "extentAndMedium", "repository",
    "archivalHistory", "acquisition", "scopeAndContent", "appraisal", "accruals",
    "arrangement", "accessConditions", "reproductionConditions", "language", "script",
    "languageNote", "physicalCharacteristics", "findingAids", "locationOfOriginals",
    "locationOfCopies", "relatedUnitsOfDescription", "publicationNote",
    "digitalObjectPath", "digitalObjectURI", "generalNote", "subjectAccessPoints",
    "placeAccessPoints", "nameAccessPoints", "genreAccessPoints", "descriptionIdentifier",
    "institutionIdentifier", "rules", "descriptionStatus", "levelOfDetail",
    "revisionHistory", "languageOfDescription", "scriptOfDescription", "sources",
    "archivistNote", "publicationStatus", "physicalObjectName", "physicalObjectLocation",
    "physicalObjectType", "alternativeIdentifiers", "alternativeIdentifierLabels",
    "eventDates", "eventTypes", "eventStartDates", "eventEndDates",
    "eventActors", "eventActorHistories", "culture",
]
# Extended columns for output files — includes audit fields not part of AtoM import
ATOM_OUTPUT_COLUMNS = ATOM_COLUMNS + ["checksumSHA256", "scanType"]

# ==================================================
# System files to ignore during scanning
# ==================================================
SYSTEM_FILES = (
    "thumbs.db",
    "desktop.ini",
    ".ds_store",
    ".spotlight-v100",
    ".trashes"
)

# ==================================================
# Run timestamp
# ==================================================
RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

# ==================================================
# Institution display map (optional)
# ==================================================
INSTITUTION_CODE_MAP = {
    "ISAM": "Iziko Museum of South Africa",
    "DNMNH": "Ditsong National Museum of Natural History",
    "ARC" : "Agricultural Research Council"
}

# ==================================================
# Audience map — derived from asset category folder name
# Add new categories here as needed
# ==================================================
AUDIENCE_MAP = {
    "digital_vouchers": "Researchers; Scientists; Public",
    "specimen_labels": "Researchers; Data curators",
    "registers":       "Archivists; Researchers",
}
AUDIENCE_FALLBACK = "Review needed"  # Fallback audience if category not in map

# ==================================================
# Description templates — derived from asset category folder name
# Placeholders: {collectionCode}, {institutionCode}
# Add new categories here as needed
# ==================================================
DESCRIPTION_TEMPLATE_MAP = {
    "digital_vouchers": "Digital voucher image of natural science specimen — {collectionCode}",
    "specimen_labels":  "Specimen label scan — {collectionCode}",
    "registers":        "Archival collection register",
}
DESCRIPTION_FALLBACK = "Review needed — {collectionCode}"

# ==================================================
# Excluded Root Folders
# ==================================================
EXCLUDED_ROOT_FOLDERS = (
    "DAMSG_output",
    "DAMSG_mapping"
)

# ==================================================
# DwC Simple Multimedia Extension — MIME type map
# ==================================================
MIME_TYPE_MAP = {
    ".tif":  "image/tiff",
    ".tiff": "image/tiff",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".pdf":  "application/pdf",
    ".csv":  "text/csv",
    ".nef":  "image/x-nikon-nef",
    ".cr2":  "image/x-canon-cr2",
    ".cr3":  "image/x-canon-cr3",
    ".arw":  "image/x-sony-arw",
    ".dng":  "image/x-adobe-dng",
    ".orf":  "image/x-olympus-orf",
    ".rw2":  "image/x-panasonic-rw2",
}

# ==================================================
# DwC Simple Multimedia Extension — type map
# Maps file extension to dc:type controlled vocabulary
# ==================================================
DWC_TYPE_MAP = {
    ".tif":  "StillImage",
    ".tiff": "StillImage",
    ".jpg":  "StillImage",
    ".jpeg": "StillImage",
    ".png":  "StillImage",
    ".nef":  "StillImage",
    ".cr2":  "StillImage",
    ".cr3":  "StillImage",
    ".arw":  "StillImage",
    ".dng":  "StillImage",
    ".orf":  "StillImage",
    ".rw2":  "StillImage",
    ".pdf":  "Text",
    ".csv":  "Text",
}

# ==================================================
# View code map
# ==================================================
VIEW_CODE_MAP = {
    "cd": "Dorsal view of specimen cranium",
    "cv": "Ventral view of specimen cranium",
    "mrl": "Right lateral view of specimen mandible",
    "sd": "Dorsal view of specimen skin",
    "mo": "Occlusional view of specimen mandible",
    "cll": "Left lateral view of specimen cranium",
    "sv": "Ventral view of specimen skin",
    "crl": "Right lateral view of specimen cranium",
    "mll": "Left lateral view of specimen mandible",
    "label": "Close-up view of specimen label"
}

# ==================================================
# Scan warning log — populated during scanning
# ==================================================
scan_warnings = []

# Detect exiftool once so we don't log a FileNotFoundError for every file
EXIFTOOL_AVAILABLE = bool(
    subprocess.run(
        ["exiftool", "-ver"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    ).returncode == 0
    if __import__("shutil").which("exiftool") else False
)
if not EXIFTOOL_AVAILABLE:
    scan_warnings.append({"level": "WARN", "file": "", "issue": "exiftool not found on PATH — date extraction will fall back to filesystem ctime"})

# Extensions that Pillow cannot open (non-image formats)
PILLOW_UNSUPPORTED = (".pdf", ".csv", ".txt", ".xml", ".mp4", ".mov", ".avi", ".wav", ".mp3")

# ==================================================
# Checksum generation for file integrity (optional)
# ==================================================
def build_extent_summary(rows):
    """Build a human-readable extent summary from a list of scanned file rows."""
    fmt_counts = {}
    total_bytes = 0
    for item in rows:
        ext = item.get("format", "").split("/")[-1].upper() or "FILE"
        fmt_counts[ext] = fmt_counts.get(ext, 0) + 1
        try:
            total_bytes += os.path.getsize(item.get("fullPath", ""))
        except Exception:
            pass
    total_mb = total_bytes / (1024 * 1024)
    size_str = f"{total_mb:.1f} MB" if total_mb >= 1 else f"{total_bytes / 1024:.1f} KB"
    fmt_str = "; ".join(f"{count} {fmt}" for fmt, count in sorted(fmt_counts.items()))
    n = len(rows)
    return f"{n} item{'s' if n != 1 else ''}: {fmt_str} ({size_str} total)"

def generate_checksum(file_path, block_size=65536):
    sha256 = hashlib.sha256()

    try:
        with open(file_path, "rb") as f:
            for block in iter(lambda: f.read(block_size), b""):
                sha256.update(block)

        return sha256.hexdigest()

    except Exception as e:
        scan_warnings.append({"level": "ERROR", "file": file_path, "issue": f"Checksum failed: {e}"})
        return ""

# ==================================================
# Hybrid date extraction
# ==================================================
def getDateCreated(path):
    raw_exts = (".nef", ".cr2", ".cr3", ".arw", ".dng", ".orf", ".rw2")
    ext = os.path.splitext(path)[1].lower()
    if EXIFTOOL_AVAILABLE:
        try:
            result = subprocess.run(
                ["exiftool", "-s3",
                 "-DateTimeOriginal", "-CreateDate", "-DateCreated",
                 "-XMP:CreateDate", "-XMP:DateCreated", path],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True
            )
            for line in result.stdout.splitlines():
                if line.strip():
                    return line.replace("-", ":")[:19]
        except Exception as e:
            scan_warnings.append({"level": "WARN", "file": path, "issue": f"exiftool error: {e}"})
    if ext not in raw_exts and ext not in PILLOW_UNSUPPORTED:
        try:
            with Image.open(path) as img:
                exif = img._getexif()
                if exif:
                    for tag_id, value in exif.items():
                        tag = ExifTags.TAGS.get(tag_id, tag_id)
                        if tag in ("DateTimeOriginal", "DateTime"):
                            return value
        except Exception as e:
            scan_warnings.append({"level": "WARN", "file": path, "issue": f"EXIF read failed: {e}"})
    try:
        ts = os.path.getctime(path)
        return datetime.fromtimestamp(ts).strftime("%Y:%m:%d %H:%M:%S")
    except Exception as e:
        scan_warnings.append({"level": "ERROR", "file": path, "issue": f"Date fallback failed: {e}"})
        return ""

# ==================================================
# Deterministic documentId generator for image/assets
# ==================================================
def generate_document_id(institution_code,collection_code, base_name, relative_path, length=8):
    relative_path = relative_path.replace("\\", "/")
    clean_inst = institution_code.replace("_", "").replace(" ", "")
    clean_collection = collection_code.replace("_", "").replace(" ", "")
    clean_base = base_name.replace("_", "").replace(" ", "")
    h = hashlib.sha1(relative_path.encode("utf-8")).hexdigest()[:length]
    return f"{clean_inst}{clean_collection}{clean_base}{h}"

# ==================================================
# Deterministic metadata documentId generator
# ==================================================
def generate_metadata_document_id(institution_code, collection_code, category, relative_path, length=8):
    relative_path = relative_path.replace("\\", "/")
    clean_inst = institution_code.replace("_", "")
    clean_collection = collection_code.replace("_", "")
    clean_category = category.replace("_", "").upper()
    h = hashlib.sha1(relative_path.encode("utf-8")).hexdigest()[:length]
    return f"{clean_inst}{clean_collection}METADATAINVENTORY{clean_category}{h}"

# ==================================================
# Tkinter UI
# ==================================================
root = Tk()
root.title("SANSCA Digital Asset Metadata Sheet Generator")
root.geometry("850x950")
root.resizable(False, False)

if not EXIFTOOL_AVAILABLE:
    messagebox.showwarning(
        "exiftool not found",
        "exiftool was not found on your PATH.\n\n"
        "Date extraction will fall back to filesystem creation time, "
        "which may not reflect the original capture date.\n\n"
        "Mac:     brew install exiftool\n"
        "Windows: download exiftool(-k).exe from exiftool.org,\n"
        "         rename to exiftool.exe, place in C:\\Windows\\"
    )

Label(root,text="Select Root Folder and Mapping CSV",font=("Arial",12,"bold")).pack(pady=10)

rootFolderVar = StringVar()
mappingFileVar = StringVar()
institutionVar = StringVar()
collectionVar = StringVar()
fileFilterVar = StringVar()
outputChoiceVar = StringVar()
scanModeVar = StringVar()
scanTypeVar = StringVar()
clearPreviousMetadataVar = BooleanVar(value=False)
clearMasterFilesVar = BooleanVar(value=False)

fileFilters = ["All","TIFF only","RAW only","JPEG only", "PDF only"]
outputChoices = ["CSV only","Excel only","Both"]
scanModes = ["Single Collection","All Collections (selected institution)","All Institutions + Collections"]
scanTypes = [
    "Working Drive",
    "Mirror Drive",
    "Mirror RAID a.k.a Suzie",
    "Collection Copy",
    "NAS Storage Repository"
] 
fileFilterVar.set(fileFilters[0])
outputChoiceVar.set(outputChoices[0])
scanModeVar.set(scanModes[0])
scanTypeVar.set(scanTypes[0])

mappingDF = None
atomMappingDF = None

# ==================================================
# UI functions
# ==================================================
script_dir = pathlib.Path(__file__).parent.resolve()
logo_path = script_dir / "nscf_logo_crop.png"

canvas = Canvas(root, width=300, height=80, bg="white", highlightthickness=0)
canvas.pack(pady=5)

try:
    logo_img = Image.open(logo_path).convert("RGB")
    logo_img = logo_img.resize((300, 80), Image.Resampling.LANCZOS)
    root.logo_tk = ImageTk.PhotoImage(logo_img)
    canvas.create_image(0, 0, anchor="nw", image=root.logo_tk)
except Exception as e:
    canvas.create_rectangle(0, 0, 300, 150, fill="lightgray")
    canvas.create_text(150, 75, text="[LOGO]", font=("Arial", 20, "bold"), fill="gray")

Label(root, text="Select Root Folder", font=("Arial", 12, "bold")).pack(pady=10)

def selectRootFolder():
    p = filedialog.askdirectory()
    if not p:
        return
    rootFolderVar.set(p)
    rootLabel.config(text=p)

def loadGoogleSheets():
    global mappingDF, atomMappingDF
    try:
        sheetStatusLabel.config(text="Connecting…", fg="gray")
        root.update()
        mappingDF     = pd.read_csv(SHEET_CSV_URL(id=SHEET_ID, sheet="master_la_collections"))
        atomMappingDF = pd.read_csv(SHEET_CSV_URL(id=SHEET_ID, sheet="master_atom_collections"))

        # Build institution name map dynamically from both sheets
        global INSTITUTION_CODE_MAP
        INSTITUTION_CODE_MAP = dict(zip(
            mappingDF["institutionCode"].dropna(),
            mappingDF["holdingInstitution"].dropna()
        ))
        # Supplement with AtoM-only institutions (institutionCode → repository)
        if "institutionCode" in atomMappingDF.columns and "repository" in atomMappingDF.columns:
            for _, row in atomMappingDF.dropna(subset=["institutionCode", "repository"]).iterrows():
                code = str(row["institutionCode"]).strip()
                name = str(row["repository"]).strip()
                if code and code not in INSTITUTION_CODE_MAP:
                    INSTITUTION_CODE_MAP[code] = name
        updateInstitutionOptions()
        updateCollectionOptions()
        updateScanModeUI()
        sheetStatusLabel.config(
            text=f"Connected — {len(mappingDF)} LA rows, {len(atomMappingDF)} AtoM rows",
            fg="green"
        )
    except Exception as e:
        messagebox.showerror("Google Sheets Error", str(e))
        sheetStatusLabel.config(text="Connection failed", fg="red")

def updateInstitutionOptions():
    if mappingDF is None:
        return
    institutions = sorted(mappingDF["institutionCode"].dropna().unique())
    institutionMenu["menu"].delete(0,"end")
    for inst in institutions:
        institutionMenu["menu"].add_command(label=inst,command=lambda v=inst: institutionVar.set(v))
    if institutions:
        institutionVar.set(institutions[0])

def updateCollectionOptions(*args):
    if mappingDF is None:
        return
    inst = institutionVar.get()
    sub = mappingDF[mappingDF["institutionCode"]==inst]
    collections = sorted(sub["collectionCode"].dropna().unique())
    collectionMenu["menu"].delete(0,"end")
    for col in collections:
        collectionMenu["menu"].add_command(label=col,command=lambda v=col: collectionVar.set(v))
    if collections:
        collectionVar.set(collections[0])

def updateScanModeUI(*args):
    mode = scanModeVar.get()
    if mode=="Single Collection":
        institutionMenu.config(state=NORMAL)
        collectionMenu.config(state=NORMAL)
    elif mode=="All Collections (selected institution)":
        institutionMenu.config(state=NORMAL)
        collectionMenu.config(state=DISABLED)
        collectionVar.set("All Collections")
    else:
        institutionMenu.config(state=DISABLED)
        collectionMenu.config(state=DISABLED)
        institutionVar.set("All Institutions")
        collectionVar.set("All Collections")

institutionVar.trace_add("write", updateCollectionOptions)
scanModeVar.trace_add("write", updateScanModeUI)


# ==================================================
# UI layout
# ==================================================
Button(root,text="Select SANSCA Root Folder",command=selectRootFolder,bg="lightgreen").pack(fill="x", padx=20, pady=5)
rootLabel=Label(root,text="",wraplength=800,anchor="w")
rootLabel.pack()

Button(root,text="Load Mapping from Google Sheets",command=loadGoogleSheets,bg="lightyellow").pack(fill="x", padx=20, pady=5)
sheetStatusLabel=Label(root,text="Mapping not loaded",wraplength=800,anchor="w",fg="gray")
sheetStatusLabel.pack()

Label(root,text="Scan Type:").pack(pady=5)
OptionMenu(root,scanTypeVar,*scanTypes).pack(fill="x", padx=20)

Label(root,text="Scan Mode:").pack(pady=5)
OptionMenu(root,scanModeVar,*scanModes).pack(fill="x", padx=20)

Label(root,text="Institution:").pack(pady=5)
institutionMenu=OptionMenu(root,institutionVar,"")
institutionMenu.pack(fill="x", padx=20)
Label(root,text="Collection:").pack(pady=5)
collectionMenu=OptionMenu(root,collectionVar,"")
collectionMenu.pack(fill="x", padx=20)

Label(root,text="File Filter:").pack(pady=5)
OptionMenu(root,fileFilterVar,*fileFilters).pack(fill="x", padx=20)

Label(root,text="Output Choice:").pack(pady=5)
OptionMenu(root,outputChoiceVar,*outputChoices).pack(fill="x", padx=20)

Checkbutton(
    root,
    text="Clear previous metadata files before scan (testing only)",
    variable=clearPreviousMetadataVar,
    fg="red"
).pack(pady=5)

Checkbutton(
    root,
    text="Clear master inventory files before scan (testing only)",
    variable=clearMasterFilesVar,
    fg="red"
).pack(pady=2)

Button(root,text="Start Processing",command=root.destroy,bg="lightblue").pack(pady=20)

root.mainloop()

# ==================================================
# Validation
# ==================================================
if mappingDF is None:
    sys.exit("Google Sheets mapping not loaded. Please select a credentials.json file.")
rootFolder = rootFolderVar.get()
institution = institutionVar.get()
collection = collectionVar.get()
scanMode = scanModeVar.get()
scanType = scanTypeVar.get()

# ==================================================
# File scanning logic
# ==================================================
fileTypes = {
    "All":[ ".tif",".tiff",".jpg",".jpeg",".nef",".cr2",".cr3",".arw",".dng",".orf",".rw2", ".pdf",".csv"],
    "TIFF only":[ ".tif",".tiff"],
    "RAW only":[ ".nef",".cr2",".cr3",".arw",".dng",".orf",".rw2"],
    "JPEG only":[ ".jpg",".jpeg"],
    "PDF only":[ ".pdf"]
}
extensions = tuple(fileTypes[fileFilterVar.get()])
all_rows = []
atom_rows = []

def scan_collection(categoryRoot, institutionCode, collectionCode, meta, atom_meta=None):
    collectionRoot = os.path.join(rootFolder, categoryRoot, institutionCode, collectionCode)
    if not os.path.isdir(collectionRoot):
        return []

    rows = []
    for r, _, files in os.walk(collectionRoot):
        for f in files: 

            # Skip hidden/system files
            if f.startswith(".") or f.startswith("._") or f.lower() in SYSTEM_FILES:
                continue
           
            if f.lower().endswith(extensions):
                full = os.path.join(r, f)
                rel = os.path.relpath(full, rootFolder)
                base = os.path.splitext(f)[0]
                fmt = os.path.splitext(f)[1].lower()

                _update_progress(f, collectionCode)
                checksum = generate_checksum(full)
                date_created = getDateCreated(full)

                asset_category = categoryRoot
                if os.path.sep + "metadata" + os.path.sep in full:
                    asset_category = f"{categoryRoot}_metadata"

                parts = base.split("_")
                view_code = parts[1] if len(parts) > 1 else ""
                view_desc = VIEW_CODE_MAP.get(view_code.lower(), "")
                if view_desc:
                    description_text = view_desc
                elif categoryRoot.lower() in DESCRIPTION_TEMPLATE_MAP:
                    description_text = DESCRIPTION_TEMPLATE_MAP[categoryRoot.lower()].format(
                        collectionCode=collectionCode,
                        institutionCode=institutionCode
                    )
                else:
                    description_text = meta.get("description", collectionCode) if meta is not None else collectionCode

                if fmt == ".csv":
                    doc_id = generate_metadata_document_id(institutionCode, collectionCode, categoryRoot, rel)
                else:
                    doc_id = generate_document_id(institutionCode, collectionCode, base, rel)

                mime_type = MIME_TYPE_MAP.get(fmt, "application/octet-stream")
                dwc_type  = DWC_TYPE_MAP.get(fmt, "")

                row_data = {
                    # --- DwC Simple Multimedia Extension standard fields ---
                    "identifier":    "",  # placeholder — URI to be assigned when image service is live
                    "type":          dwc_type,
                    "format":        mime_type,
                    "title":         base,
                    "description":   description_text,
                    "created":       date_created,
                    "creator":       meta.get("creator", "") if meta is not None else "",
                    "contributor":   meta.get("contributor", "") if meta is not None else "",
                    "publisher":     meta.get("publisher", "") if meta is not None else "",
                    "audience":      AUDIENCE_MAP.get(categoryRoot.lower(), AUDIENCE_FALLBACK),
                    "source":        meta.get("source", "") if meta is not None else "",
                    "license":       meta.get("license", "") if meta is not None else "",
                    "rightsHolder":  meta.get("rightsHolder", "") if meta is not None else "",
                    "references":    "",  # placeholder — URI to occurrence record, future field
                    # --- System / archival fields ---
                    "scanType":        scanType,
                    "documentId":      doc_id,
                    "institutionCode": institutionCode,
                    "collectionCode":  collectionCode,
                    "institutionName":    INSTITUTION_CODE_MAP.get(institutionCode, institutionCode),
                    "holdingInstitution": (
                        meta.get("holdingInstitution", "") if meta is not None
                        else (atom_meta.get("repository", "") if atom_meta is not None and len(atom_meta) > 0 else INSTITUTION_CODE_MAP.get(institutionCode, institutionCode))
                    ),
                    "additionalNames": "",
                    "subject":         "Metadata" if fmt == ".csv" else (meta.get("subject", "") if meta is not None else ""),
                    "fileName":        f,
                    "fullPath":        full,
                    "relativePath":    rel,
                    "assetCategory":   asset_category,
                    "scanModeApplied": scanMode,
                    "checksumSHA256":  checksum,
                    # preserved for legacy/archival use
                    "dateCreated":     date_created,
                }

                # Add ALL mapping columns automatically
                if meta is not None and mappingDF is not None:
                    for col in mappingDF.columns:
                        if col not in row_data or (col == "additionalNames" and not row_data["additionalNames"].strip()):
                            row_data[col] = meta.get(col, "")

                rows.append(row_data)   

    return rows

# ==================================================
# Master file paths
# ==================================================
master_csv = os.path.join(rootFolder, "DAMSG_output", f"digital_asset_inventory_la_{RUN_TIMESTAMP}.csv")
master_xlsx = os.path.join(rootFolder, "DAMSG_output", f"digital_asset_inventory_la_{RUN_TIMESTAMP}.xlsx")
os.makedirs(os.path.dirname(master_csv), exist_ok=True)

if clearMasterFilesVar.get():
    output_folder = os.path.dirname(master_csv)
    for f in os.listdir(output_folder):
        if f.startswith("digital_asset_inventory_la_") and f.endswith((".csv", ".xlsx")):
            try:
                os.remove(os.path.join(output_folder, f))
                print(f"Cleared master file: {f}")
            except Exception as e:
                print(f"Could not delete {f}: {e}")
    for f in os.listdir(output_folder):
        if f.startswith("preservation_audit_la_") and f.endswith(".csv"):
            try:
                os.remove(os.path.join(output_folder, f))
                print(f"Cleared audit file: {f}")
            except Exception as e:
                print(f"Could not delete {f}: {e}")
        if f.startswith("scan_warnings_") and f.endswith(".csv"):
            try:
                os.remove(os.path.join(output_folder, f))
                print(f"Cleared scan warnings: {f}")
            except Exception as e:
                print(f"Could not delete {f}: {e}")
        if (f.startswith("digital_asset_inventory_atom_") or f.startswith("atom_import_")) and f.endswith(".csv"):
            try:
                os.remove(os.path.join(output_folder, f))
                print(f"Cleared AtoM file: {f}")
            except Exception as e:
                print(f"Could not delete {f}: {e}")
        if f.startswith("preservation_audit_atom_") and f.endswith(".csv"):
            try:
                os.remove(os.path.join(output_folder, f))
                print(f"Cleared AtoM audit file: {f}")
            except Exception as e:
                print(f"Could not delete {f}: {e}")

output_folder = os.path.dirname(master_csv)
previous_la_files = sorted(
    [f for f in os.listdir(output_folder) if f.startswith("digital_asset_inventory_la_") and f.endswith(".csv")],
    reverse=True
)
if previous_la_files:
    master_df = pd.read_csv(os.path.join(output_folder, previous_la_files[0]))
else:
    master_df = pd.DataFrame()

previous_atom_files = sorted(
    [f for f in os.listdir(output_folder) if f.startswith("digital_asset_inventory_atom_") and f.endswith(".csv")],
    reverse=True
)
if previous_atom_files:
    atom_master_df = pd.read_csv(os.path.join(output_folder, previous_atom_files[0]))
else:
    atom_master_df = pd.DataFrame()

# ==================================================
# Progress window
# ==================================================
progress_win = Tk()
progress_win.title("Scanning…")
progress_win.resizable(False, False)
progress_win.geometry("520x90")

_prog_label = Label(progress_win, text="Starting scan…", anchor="w", padx=12, pady=8)
_prog_label.pack(fill="x")
_prog_count = Label(progress_win, text="", anchor="w", padx=12, fg="#555")
_prog_count.pack(fill="x")

_scanned_total = [0]

def _update_progress(filename, collection):
    _scanned_total[0] += 1
    _prog_label.config(text=f"[{collection}]  {filename}")
    _prog_count.config(text=f"{_scanned_total[0]} file(s) scanned")
    progress_win.update()

# ==================================================
# Scan, generate subset CSVs, and append newest metadata
# ==================================================
categories = [
    d for d in os.listdir(rootFolder)
    if os.path.isdir(os.path.join(rootFolder, d))
    and d not in EXCLUDED_ROOT_FOLDERS
]

for cat in categories:
    cat_path = os.path.join(rootFolder, cat)
    if not os.path.isdir(cat_path):
        continue

    insts = [d for d in os.listdir(cat_path) if os.path.isdir(os.path.join(cat_path, d))]
    for inst in insts:
        inst_path = os.path.join(cat_path, inst)
        collections = [d for d in os.listdir(inst_path) if os.path.isdir(os.path.join(inst_path, d))]
        for coll in collections:
            if scanMode == "Single Collection" and (inst != institution or coll != collection):
                continue
            if scanMode == "All Collections (selected institution)" and inst != institution:
                continue

            targets = CATEGORY_TARGETS.get(cat, [])
            la_rows = mappingDF[(mappingDF["institutionCode"] == inst) & (mappingDF["collectionCode"] == coll)] if mappingDF is not None else pd.DataFrame()
            if "LA" in targets and la_rows.empty:
                print(f"Skipping {inst}/{coll} — no LA mapping found")
                continue
            meta = la_rows.iloc[0] if not la_rows.empty else None

            if clearPreviousMetadataVar.get():
                meta_folder_pre = os.path.join(inst_path, coll, "metadata")
                if os.path.isdir(meta_folder_pre):
                    for old_file in os.listdir(meta_folder_pre):
                        if old_file.lower().endswith(".csv"):
                            try:
                                os.remove(os.path.join(meta_folder_pre, old_file))
                            except Exception as e:
                                print(f"Could not delete {old_file}: {e}")

            # Look up AtoM mapping row early so scan_collection can use it
            atom_meta_pre = {}
            if atomMappingDF is not None and "institutionCode" in atomMappingDF.columns and "collectionCode" in atomMappingDF.columns:
                atom_rows_match = atomMappingDF[
                    (atomMappingDF["institutionCode"] == inst) &
                    (atomMappingDF["collectionCode"] == coll)
                ]
                if not atom_rows_match.empty:
                    atom_meta_pre = atom_rows_match.iloc[0]

            scanned = scan_collection(cat, inst, coll, meta, atom_meta_pre)

            la_only = "LA" in targets
            atom_only = targets == ["AtoM"]

            if la_only or not atom_only:
                all_rows.extend(scanned)

            subset_rows = [
                r for r in scanned
                    if (
                        r["format"] != "text/csv" and
                        r["scanType"] == scanType
                    )
            ]
            if subset_rows:
                meta_folder = os.path.join(inst_path, coll, "metadata")
                os.makedirs(meta_folder, exist_ok=True)
                scanDateHuman = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                def write_header_csv(path, df):
                    with open(path, "w", newline="", encoding="utf-8") as f:
                        f.write(f"scanType,{scanType}\n")
                        f.write(f"scanMode,{scanMode}\n")
                        f.write(f"scanTimestamp,{RUN_TIMESTAMP}\n")
                        f.write(f"scanDate,{scanDateHuman}\n")
                        f.write(f"institutionCode,{inst}\n")
                        f.write(f"collectionCode,{coll}\n\n")
                        df.to_csv(f, index=False)
                    print(f"Collection metadata CSV generated: {path}")

                # LA format CSV
                if "LA" in targets:
                    subset_path = os.path.join(meta_folder, f"{coll}_{cat}_metadata_la_{RUN_TIMESTAMP}.csv")
                    write_header_csv(subset_path, pd.DataFrame(subset_rows))

                # AtoM format CSV — written after AtoM row generation below

                # Use LA path for the master row reference (LA only)
                subset_path = os.path.join(meta_folder, f"{coll}_{cat}_metadata_la_{RUN_TIMESTAMP}.csv") if "LA" in targets else ""

                now_ts = datetime.now().strftime("%Y:%m:%d %H:%M:%S")
                if meta is not None:
                    row_data = {
                        # DwC Simple Multimedia Extension standard fields
                        "identifier":    "",
                        "type":          "Text",
                        "format":        "text/csv",
                        "title":         os.path.splitext(os.path.basename(subset_path))[0],
                        "description":   (
                            f"Metadata file for {inst}_{coll}" if scanMode == "Single Collection"
                            else f"Metadata file for {inst}" if scanMode == "All Collections (selected institution)"
                            else "Metadata file for all collections held by NSCF partner institutions"
                        ) + f" [{build_extent_summary(subset_rows)}]",
                        "created":       now_ts,
                        "creator":       meta.get("creator", ""),
                        "contributor":   meta.get("contributor", ""),
                        "publisher":     meta.get("publisher", ""),
                        "audience":      "Data curators; Collection managers",
                        "source":        f"Original digital assets — {inst}_{coll} ({cat})",
                        "license":       meta.get("license", ""),
                        "rightsHolder":  meta.get("rightsHolder", ""),
                        "references":    "",
                        # System / archival fields
                        "fileName":          os.path.basename(subset_path),
                        "scanType":          scanType,
                        "documentId":        generate_metadata_document_id(inst, coll, cat, os.path.relpath(subset_path, rootFolder)),
                        "institutionCode":   inst,
                        "collectionCode":    coll,
                        "institutionName":   INSTITUTION_CODE_MAP.get(inst, inst),
                        "holdingInstitution": meta.get("holdingInstitution", ""),
                        "dateCreated":       now_ts,
                        "subject":           "Metadata",
                        "fullPath":          subset_path,
                        "relativePath":      os.path.relpath(subset_path, rootFolder),
                        "assetCategory":     f"{cat}_metadata",
                        "scanModeApplied":   scanMode,
                        "additionalNames":   "",
                        "checksumSHA256":    generate_checksum(subset_path),
                    }
                    # Add ALL mapping columns automatically (same as scan_collection)
                    for col in mappingDF.columns:
                        if col not in row_data or (col == "additionalNames" and not row_data["additionalNames"].strip()):
                            row_data[col] = meta.get(col, "")
                    all_rows.append(row_data)

            # --------------------------------------------------
            # AtoM output — generate parent + item rows
            # --------------------------------------------------
            if "AtoM" in CATEGORY_TARGETS.get(cat, []) and atomMappingDF is not None:
                try:
                    atom_meta = atomMappingDF[
                        (atomMappingDF["institutionCode"] == inst) &
                        (atomMappingDF["collectionCode"] == coll)
                    ].iloc[0]
                except IndexError:
                    scan_warnings.append({"level": "WARN", "file": "", "issue": f"No AtoM mapping row for {inst}/{coll} — skipped"})
                    atom_meta = {}

                parent_legacy_id = f"{inst}_{coll}_{cat}"
                inst_name = INSTITUTION_CODE_MAP.get(inst, inst)

                # Parent row
                parent_row = {col: "" for col in ATOM_COLUMNS}
                parent_row["legacyId"]           = parent_legacy_id
                parent_row["title"]              = atom_meta.get("title", f"{inst} {coll}")
                parent_row["levelOfDescription"] = atom_meta.get("levelOfDescription", "Collection")
                parent_row["institutionIdentifier"] = inst
                for col in ATOM_COLUMNS:
                    if col in atom_meta and not parent_row[col]:
                        parent_row[col] = atom_meta.get(col, "")
                if not parent_row["repository"]:
                    parent_row["repository"] = inst_name

                # Build extentAndMedium summary from child items
                parent_row["extentAndMedium"] = build_extent_summary(subset_rows)

                atom_rows.append(parent_row)

                # Item rows — one per scanned file in this collection
                for item in subset_rows:
                    item_row = {col: "" for col in ATOM_COLUMNS}
                    item_row["parentId"]            = parent_legacy_id
                    item_row["identifier"]          = item.get("documentId", "").replace(" ", "_")
                    item_row["title"]               = item.get("title", "")
                    item_row["levelOfDescription"]  = "Item"
                    item_row["repository"]          = atom_meta.get("repository", "") or inst_name
                    item_row["institutionIdentifier"] = inst
                    item_row["digitalObjectPath"]   = item.get("relativePath", "")
                    item_row["eventDates"]          = item.get("dateCreated", "")
                    item_row["eventStartDates"]     = item.get("dateCreated", "")
                    item_row["eventTypes"]          = atom_meta.get("eventTypes", "creation")
                    item_row["eventActors"]         = atom_meta.get("eventActors", "")
                    item_row["eventActorHistories"] = atom_meta.get("eventActorHistories", "")
                    item_row["language"]            = atom_meta.get("language", "")
                    item_row["script"]              = atom_meta.get("script", "")
                    item_row["accessConditions"]    = atom_meta.get("accessConditions", "")
                    item_row["reproductionConditions"] = atom_meta.get("reproductionConditions", "")
                    item_row["publicationStatus"]   = atom_meta.get("publicationStatus", "")
                    item_row["culture"]             = atom_meta.get("culture", "")
                    item_row["extentAndMedium"]     = f"1 {item.get('format', '').split('/')[-1].upper()} file"
                    item_row["checksumSHA256"]      = item.get("checksumSHA256", "")
                    item_row["scanType"]            = item.get("scanType", "")
                    atom_rows.append(item_row)

                # Write AtoM per-collection metadata CSV now that rows are generated
                if subset_rows:
                    meta_folder_atom = os.path.join(inst_path, coll, "metadata")
                    os.makedirs(meta_folder_atom, exist_ok=True)
                    atom_subset_rows = [r for r in atom_rows if str(r.get("parentId", "") or "").startswith(f"{inst}_{coll}_")]
                    if atom_subset_rows:
                        atom_subset_path = os.path.join(meta_folder_atom, f"{coll}_{cat}_metadata_atom_{RUN_TIMESTAMP}.csv")
                        scanDateHuman = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        with open(atom_subset_path, "w", newline="", encoding="utf-8") as f:
                            f.write(f"scanType,{scanType}\n")
                            f.write(f"scanMode,{scanMode}\n")
                            f.write(f"scanTimestamp,{RUN_TIMESTAMP}\n")
                            f.write(f"scanDate,{scanDateHuman}\n")
                            f.write(f"institutionCode,{inst}\n")
                            f.write(f"collectionCode,{coll}\n\n")
                            pd.DataFrame(atom_subset_rows, columns=ATOM_OUTPUT_COLUMNS).to_csv(f, index=False)
                        print(f"Collection metadata CSV generated: {atom_subset_path}")


progress_win.destroy()

# ==================================================
# Convert all_rows to DataFrame safely
# ==================================================

expected_columns = [
    # DwC Simple Multimedia Extension standard fields (always present)
    "identifier", "type", "format", "title", "description",
    "created", "creator", "contributor", "publisher",
    "audience", "source", "license", "rightsHolder", "references",
    # System / archival fields
    "scanType", "documentId", "fileName", "relativePath", "fullPath",
    "assetCategory", "dateCreated", "scanModeApplied",
    "institutionCode", "collectionCode", "institutionName",
    "additionalNames", "holdingInstitution", "subject", "checksumSHA256",
]

for col in expected_columns:
    if col not in master_df.columns:
        master_df[col] = ""

# Create new rows dataframe
if all_rows:
    new_rows_df = pd.DataFrame(all_rows)
else:
    new_rows_df = pd.DataFrame(columns=expected_columns)

# Ensure master_df exists and has expected structure
if master_df.empty:
    master_df = pd.DataFrame(columns=expected_columns)

# Ensure documentId column exists in both
if "documentId" not in new_rows_df.columns:
    new_rows_df["documentId"] = ""

if "documentId" not in master_df.columns:
    master_df["documentId"] = ""

# Remove duplicates safely
if not new_rows_df.empty:
    new_rows_df = new_rows_df[
    ~new_rows_df.set_index(["documentId","scanType"]).index.isin(
        master_df.set_index(["documentId","scanType"]).index
    )
]

# Preserve description for CSV metadata safely
required_cols = {"format", "description", "documentId", "institutionCode", "collectionCode"}
if not new_rows_df.empty and required_cols.issubset(new_rows_df.columns):
    mask_csv = new_rows_df["format"] == "text/csv"

    if not master_df.empty and "documentId" in master_df.columns:
        existing_desc = master_df.set_index("documentId")["description"].to_dict()

        def preserve_or_generate_desc(row):
            doc_id = row["documentId"]
            if doc_id in existing_desc and existing_desc[doc_id].strip():
                return existing_desc[doc_id]
            return row.get("description", "")

        new_rows_df.loc[mask_csv, "description"] = new_rows_df.loc[mask_csv].apply(preserve_or_generate_desc, axis=1)
    else:
        pass  # description already set on row during scan

# Fill additionalNames if empty
if not new_rows_df.empty and "additionalNames" in new_rows_df.columns and "additionalNames" in mappingDF.columns:
    mask_additional = new_rows_df["additionalNames"].isna() | (new_rows_df["additionalNames"].str.strip() == "")
    def fill_additional_names(row):
        df_match = mappingDF[
            (mappingDF["institutionCode"] == row["institutionCode"]) &
            (mappingDF["collectionCode"] == row["collectionCode"])
        ]
        if not df_match.empty:
            return df_match.iloc[0].get("additionalNames", "")
        return ""
    
    new_rows_df.loc[mask_additional, "additionalNames"] = (
    new_rows_df.loc[mask_additional]
    .apply(fill_additional_names, axis=1)
)

# ==================================================
# Append new rows to master
# ==================================================
updated_master_df = pd.concat([master_df, new_rows_df], ignore_index=True)

# Preserve column order
mapping_columns = [col for col in mappingDF.columns if col in updated_master_df.columns]
system_columns = [col for col in [
    "scanType","documentId","title","fileName","relativePath","fullPath",
    "format","assetCategory","dateCreated","scanModeApplied","checksumSHA256"
] if col in updated_master_df.columns]

ordered_columns = system_columns + [col for col in mapping_columns if col not in system_columns]
updated_master_df = updated_master_df[ordered_columns]

# ==================================================
# Ensure checksum column exists before preservation audit
# ==================================================
if "checksumSHA256" not in updated_master_df.columns:
    updated_master_df["checksumSHA256"] = ""

# ==================================================
# Preservation Audit Report (Revised)
# ==================================================
from collections import Counter

def run_preservation_audit(master_df, atom_df=None):

    audit_folder = os.path.join(rootFolder, "DAMSG_output")
    os.makedirs(audit_folder, exist_ok=True)

    audit_base      = os.path.join(audit_folder, f"preservation_audit_la_{RUN_TIMESTAMP}")
    atom_audit_base = os.path.join(audit_folder, f"preservation_audit_atom_{RUN_TIMESTAMP}")

    # Only evaluate real files (skip .csv metadata)
    df = master_df[master_df["format"] != "text/csv"].copy()

    # Define only logical source→target storage comparisons
    allowed_pairs = [
        ("Working Drive", "Mirror Drive"),
        ("Mirror Drive", "Mirror RAID a.k.a Suzie"),
        ("Working Drive", "Collection Copy"),
        ("Mirror RAID a.k.a Suzie", "NAS Storage Repository")
    ]

    summary_rows = []
    missing_rows = []
    mismatch_rows = []
    duplicate_rows = []

    for source, target in allowed_pairs:

        source_df = df[df["scanType"] == source]
        target_df = df[df["scanType"] == target]

        source_index = dict(zip(source_df["relativePath"], source_df["checksumSHA256"]))
        target_index = dict(zip(target_df["relativePath"], target_df["checksumSHA256"]))

        missing = []
        mismatch = []
        matching = []

        for path, checksum in source_index.items():
            if path not in target_index:
                missing.append(path)
            elif target_index[path] != checksum:
                mismatch.append(path)
            else:
                matching.append(path)

        summary_rows.append({
            "Source Storage": source,
            "Target Storage": target,
            "Total Source Files": len(source_index),
            "Matching": len(matching),
            "Missing on Target": len(missing),
            "Checksum Mismatch": len(mismatch)
        })

        for p in missing:
            missing_rows.append({
                "Source Storage": source,
                "Missing From": target,
                "relativePath": p
            })

        for p in mismatch:
            mismatch_rows.append({
                "Source Storage": source,
                "Mismatch With": target,
                "relativePath": p
            })

    # Detect duplicates across all files
    checksum_counts = Counter(df["checksumSHA256"])
    duplicate_checksums = [k for k, v in checksum_counts.items() if v > 1]

    for chk in duplicate_checksums:
        dup_files = df[df["checksumSHA256"] == chk]
        for _, row in dup_files.iterrows():
            duplicate_rows.append({
                "checksum": chk,
                "relativePath": row["relativePath"],
                "scanType": row["scanType"]
            })

    # Create DataFrames
    summary_df = pd.DataFrame(summary_rows)
    missing_df = pd.DataFrame(missing_rows)
    mismatch_df = pd.DataFrame(mismatch_rows)
    duplicates_df = pd.DataFrame(duplicate_rows)

    # AtoM coverage audit — same cross-storage logic as LA
    atom_missing_rows = []
    atom_mismatch_rows = []
    atom_summary_rows = []
    atom_duplicate_rows = []
    if atom_df is not None and len(atom_df) > 0:
        atom_items = pd.DataFrame([r for r in atom_df if r.get("levelOfDescription") == "Item"])
        if not atom_items.empty:
            atom_items = atom_items.rename(columns={"digitalObjectPath": "relativePath"})

            for source, target in allowed_pairs:
                source_df = atom_items[atom_items["scanType"] == source]
                target_df = atom_items[atom_items["scanType"] == target]

                source_index = dict(zip(source_df["relativePath"], source_df["checksumSHA256"]))
                target_index = dict(zip(target_df["relativePath"], target_df["checksumSHA256"]))

                missing, mismatch, matching = [], [], []
                for path, checksum in source_index.items():
                    if path not in target_index:
                        missing.append(path)
                    elif target_index[path] != checksum:
                        mismatch.append(path)
                    else:
                        matching.append(path)

                atom_summary_rows.append({
                    "Source Storage": source, "Target Storage": target,
                    "Total Source Files": len(source_index),
                    "Matching": len(matching),
                    "Missing on Target": len(missing),
                    "Checksum Mismatch": len(mismatch)
                })
                for p in missing:
                    atom_missing_rows.append({"Source Storage": source, "Missing From": target, "relativePath": p})
                for p in mismatch:
                    atom_mismatch_rows.append({"Source Storage": source, "Mismatch With": target, "relativePath": p})

            checksum_counts = Counter(atom_items["checksumSHA256"])
            for chk, count in checksum_counts.items():
                if count > 1:
                    for _, row in atom_items[atom_items["checksumSHA256"] == chk].iterrows():
                        atom_duplicate_rows.append({"checksum": chk, "relativePath": row["relativePath"], "scanType": row["scanType"]})

    # Write CSVs — one per section, only if non-empty
    audit_file = f"{audit_base}_summary.csv"
    summary_df.to_csv(audit_file, index=False)

    if not missing_df.empty:
        missing_df.to_csv(f"{audit_base}_missing.csv", index=False)
    if not mismatch_df.empty:
        mismatch_df.to_csv(f"{audit_base}_mismatch.csv", index=False)
    if not duplicates_df.empty:
        duplicates_df.to_csv(f"{audit_base}_duplicates.csv", index=False)
    if atom_summary_rows:
        pd.DataFrame(atom_summary_rows).to_csv(f"{atom_audit_base}_summary.csv", index=False)
    if atom_missing_rows:
        pd.DataFrame(atom_missing_rows).to_csv(f"{atom_audit_base}_missing.csv", index=False)
    if atom_mismatch_rows:
        pd.DataFrame(atom_mismatch_rows).to_csv(f"{atom_audit_base}_mismatch.csv", index=False)
    if atom_duplicate_rows:
        pd.DataFrame(atom_duplicate_rows).to_csv(f"{atom_audit_base}_duplicates.csv", index=False)

    print(f"Preservation audit report created: {audit_file}")

    return audit_file

# ==================================================
# Write master CSV/Excel
# ==================================================
output_choice = outputChoiceVar.get()
if output_choice in ("CSV only", "Both"):
    updated_master_df.to_csv(master_csv, index=False)
    print(f"Processing complete. Master CSV updated: {master_csv}")
if output_choice in ("Excel only", "Both"):
    updated_master_df.to_excel(master_xlsx, index=False)
    print(f"Processing complete. Master Excel updated: {master_xlsx}")

# ==================================================
# Write scan warning log
# ==================================================
if scan_warnings:
    log_path = os.path.join(rootFolder, "DAMSG_output", f"scan_warnings_{RUN_TIMESTAMP}.csv")
    pd.DataFrame(scan_warnings).to_csv(log_path, index=False)
    print(f"Scan warnings written ({len(scan_warnings)} issues): {log_path}")
else:
    print("Scan completed with no warnings.")

# ==================================================
# Accumulate AtoM master (same pattern as LA)
# ==================================================
atom_csv = None
if atom_rows:
    new_atom_df = pd.DataFrame(atom_rows, columns=ATOM_OUTPUT_COLUMNS)
    updated_atom_df = pd.concat([atom_master_df, new_atom_df], ignore_index=True)
    # Deduplicate: use digitalObjectPath+scanType for items, legacyId for parents
    if "digitalObjectPath" in updated_atom_df.columns and "scanType" in updated_atom_df.columns:
        updated_atom_df = updated_atom_df.reset_index(drop=True)
        item_mask = updated_atom_df["levelOfDescription"] == "Item"
        parent_mask = ~item_mask
        items_deduped = updated_atom_df[item_mask].drop_duplicates(subset=["digitalObjectPath", "scanType"], keep="last")
        parents_deduped = updated_atom_df[parent_mask].drop_duplicates(subset=["legacyId"], keep="last")
        # Rebuild interleaved order: parent then its items, in original row order
        all_deduped = pd.concat([parents_deduped, items_deduped]).sort_index()
        updated_atom_df = all_deduped.reset_index(drop=True)
    atom_csv = os.path.join(rootFolder, "DAMSG_output", f"digital_asset_inventory_atom_{RUN_TIMESTAMP}.csv")
    updated_atom_df.to_csv(atom_csv, index=False)
    print(f"AtoM master CSV written ({len(updated_atom_df)} rows): {atom_csv}")
else:
    updated_atom_df = atom_master_df

# ==================================================
# Run preservation audit
# ==================================================
audit_path = run_preservation_audit(
    updated_master_df,
    updated_atom_df.to_dict("records") if not updated_atom_df.empty else None
)

# ==================================================
# Optionally, open files automatically
# ==================================================
def open_file(filepath):
    system = platform.system()
    try:
        if system=="Windows":
            os.startfile(filepath)
        elif system=="Darwin":
            subprocess.run(["open",filepath],check=True)
        else:
            subprocess.run(["xdg-open",filepath],check=True)
    except Exception as e:
        print(f"Could not open {filepath}: {e}")

damsg_output_folder = os.path.join(rootFolder, "DAMSG_output")
for f in sorted(os.listdir(damsg_output_folder)):
    if f.endswith(".csv"):
        open_file(os.path.join(damsg_output_folder, f))
if outputChoiceVar.get() in ("Both", "Excel only"):
    open_file(master_xlsx)