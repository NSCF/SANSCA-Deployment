#!/usr/bin/env python3
import os
import subprocess
import pandas as pd
from datetime import datetime
from tkinter import (
    Tk, Label, Button, StringVar, OptionMenu,
    filedialog, messagebox, DISABLED, NORMAL
)
from PIL import Image, ExifTags, ImageTk
import platform
import sys

# ==================================================
# Run timestamp
# ==================================================
RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

# ==================================================
# Institution display map (optional)
# ==================================================
INSTITUTION_CODE_MAP = {
    "IZIKO": "Iziko Museum of South Africa",
    "DNMNH": "Ditsong National Museum of Natural History"
}

# ==================================================
# Hybrid date extraction
# ==================================================
def getDateCreated(path):
    raw_exts = (".nef", ".cr2", ".cr3", ".arw", ".dng", ".orf", ".rw2")
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
    except Exception:
        pass
    if not path.lower().endswith(raw_exts):
        try:
            with Image.open(path) as img:
                exif = img._getexif()
                if exif:
                    for tag_id, value in exif.items():
                        tag = ExifTags.TAGS.get(tag_id, tag_id)
                        if tag in ("DateTimeOriginal", "DateTime"):
                            return value
        except Exception:
            pass
    try:
        ts = os.path.getctime(path)
        return datetime.fromtimestamp(ts).strftime("%Y:%m:%d %H:%M:%S")
    except Exception:
        return ""

# ==================================================
# Tkinter UI
# ==================================================
root = Tk()
root.title("SANSCA Digital Asset Metadata Sheet Generator")
root.geometry("850x800")

Label(root,text="Select Root Folder and Mapping CSV",font=("Arial",12,"bold")).pack(pady=10)

rootFolderVar = StringVar()
mappingFileVar = StringVar()
institutionVar = StringVar()
collectionVar = StringVar()
fileFilterVar = StringVar()
outputChoiceVar = StringVar()
languageVar = StringVar()
scanModeVar = StringVar()

fileFilters = ["All","TIFF only","RAW only","JPEG only"]
outputChoices = ["CSV only","Excel only","Both"]
languages = ["English","Afrikaans","Latin/Greek"]
scanModes = ["Single Collection","All Collections (selected institution)","All Institutions + Collections"]

fileFilterVar.set(fileFilters[0])
outputChoiceVar.set(outputChoices[0])
languageVar.set(languages[0])
scanModeVar.set(scanModes[0])

mappingDF = None

# ==================================================
# UI functions
# ==================================================
#/Users/klippies/Projects/Coding/github_repository_clones/SANSCA-Deployment/tools/data_manipulation/Windows/NSCF-Logo-crop.png
# Logo
logo_path = os.path.join(
    os.path.dirname(__file__), "NSCF-Logo-crop.png"
)

try:
    logo_img = Image.open(logo_path)
    # Resize to fit nicely at the top
    logo_img = logo_img.resize((300, 150), Image.Resampling.LANCZOS)
    root.logo_tk = ImageTk.PhotoImage(logo_img)
    logo_label = Label(root, image=root.logo_tk)
    logo_label.pack(pady=10)
except Exception as e:
    # fallback placeholder if logo can't be loaded
    Label(root, text="[LOGO]", font=("Arial", 20, "bold"), fg="gray").pack(pady=10)
    print(f"Logo could not be loaded: {e}")

# -------------------------------
# Main heading
# -------------------------------
Label(root, text="Select Root Folder", font=("Arial", 12, "bold")).pack(pady=10)

def selectRootFolder():
    p = filedialog.askdirectory()
    if not p:
        return
    rootFolderVar.set(p)
    rootLabel.config(text=p)

    # --------------------------------------------------
    # Auto-load mapping from SANSCA/DAMSG_mapping
    # --------------------------------------------------
    global mappingDF
    mapping_path = os.path.join(p, "DAMSG_mapping", "collections_mapping_2026.csv")
    if not os.path.isfile(mapping_path):
        messagebox.showerror("Mapping CSV Error",
                             f"Mapping file not found at:\n{mapping_path}")
        mappingDF = None
        return

    try:
        df = pd.read_csv(mapping_path)
    except Exception as e:
        messagebox.showerror("CSV Error", str(e))
        mappingDF = None
        return

    # Check required columns
    required = {"institutionCode","collectionCode","creator","contributor",
                "license","rightsHolder","holdingInstitution","description"}
    missing = required - set(df.columns)
    if missing:
        messagebox.showerror("Mapping CSV Error",
                             f"Missing required columns:\n{', '.join(sorted(missing))}")
        mappingDF = None
        return

    mappingDF = df
    # Update institution and collection options
    updateInstitutionOptions()
    updateCollectionOptions()
    updateScanModeUI()

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
    else:  # All Institutions + Collections
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
#Button(root,text="Select Metadata Mapping CSV",command=selectMappingCSV,bg="orange").pack(fill="x", padx=20, pady=5)
#mappingLabel=Label(root,text="",wraplength=800,anchor="w")
#mappingLabel.pack()

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
Label(root,text="Language:").pack(pady=5)
OptionMenu(root,languageVar,*languages).pack(fill="x", padx=20)
Label(root,text="Output Choice:").pack(pady=5)
OptionMenu(root,outputChoiceVar,*outputChoices).pack(fill="x", padx=20)

Button(root,text="Start Processing",command=root.destroy,bg="lightblue").pack(pady=20)

root.mainloop()

# ==================================================
# Validation
# ==================================================
if mappingDF is None:
    sys.exit("No mapping CSV loaded.")
rootFolder = rootFolderVar.get()
institution = institutionVar.get()
collection = collectionVar.get()
scanMode = scanModeVar.get()

# ==================================================
# File scanning logic with multi-mode
# ==================================================
fileTypes = {
    "All":[ ".tif",".tiff",".jpg",".jpeg",".nef",".cr2",".cr3",".arw",".dng",".orf",".rw2",".csv"],
    "TIFF only":[ ".tif",".tiff"],
    "RAW only":[ ".nef",".cr2",".cr3",".arw",".dng",".orf",".rw2"],
    "JPEG only":[ ".jpg",".jpeg"]
}
extensions = tuple(fileTypes[fileFilterVar.get()])
all_rows = []

def scan_collection(categoryRoot, institutionCode, collectionCode, meta):
    collectionRoot = os.path.join(rootFolder, categoryRoot, institutionCode, collectionCode)
    if not os.path.isdir(collectionRoot):
        return []

    rows = []
    for r, _, files in os.walk(collectionRoot):
        for f in files:
            if f.lower().endswith(extensions):
                full = os.path.join(r, f)
                rel = os.path.relpath(full, rootFolder)
                base = os.path.splitext(f)[0]
                fmt = os.path.splitext(f)[1].lower()

                # --------------------------------------------------
                # Determine correct asset category
                # Files inside /metadata/ get <category>_metadata
                # --------------------------------------------------
                asset_category = categoryRoot
                if os.path.sep + "metadata" + os.path.sep in full:
                    asset_category = f"{categoryRoot}_metadata"

                rows.append({
                    "documentId": f"{base}_{collectionCode}_{hash(rel) & 0xffff}",
                    "title": base,
                    "institutionCode": institutionCode,
                    "collectionCode": collectionCode,
                    "institutionName": INSTITUTION_CODE_MAP.get(institutionCode, institutionCode),
                    "creator": meta["creator"],
                    "contributor": meta["contributor"],
                    "description": f'{meta["description"]} ({collectionCode})',
                    "rightsHolder": meta["rightsHolder"],
                    "holdingInstitution": meta["holdingInstitution"],
                    "license": meta["license"],
                    "dateCreated": getDateCreated(full),
                    "format": fmt,
                    "subject": "metadata" if fmt == ".csv" else meta.get("subject", ""),
                    "language": languageVar.get(),
                    "fileName": f,
                    "fullPath": full,
                    "relativePath": rel,
                    "assetCategory": asset_category
                })

    return rows

categories = [d for d in os.listdir(rootFolder) if os.path.isdir(os.path.join(rootFolder, d))]

# ==================================================
# Scan and build main DataFrame
# ==================================================
for cat in categories:
    cat_path=os.path.join(rootFolder,cat)
    if not os.path.isdir(cat_path):
        continue
    insts=[d for d in os.listdir(cat_path) if os.path.isdir(os.path.join(cat_path,d))]
    for inst in insts:
        inst_path=os.path.join(cat_path,inst)
        collections=[d for d in os.listdir(inst_path) if os.path.isdir(os.path.join(inst_path,d))]
        for coll in collections:
            if scanMode=="Single Collection" and (inst!=institution or coll!=collection):
                continue
            if scanMode=="All Collections (selected institution)" and inst!=institution:
                continue
            try:
                meta=mappingDF[(mappingDF["institutionCode"]==inst)&(mappingDF["collectionCode"]==coll)].iloc[0]
            except IndexError:
                continue
            all_rows.extend(scan_collection(cat,inst,coll,meta))

df=pd.DataFrame(all_rows)
if df.empty:
    sys.exit("No matching files found for selected filter.")

# ==================================================
# Generate subset CSVs (images only) per collection
# ==================================================
subset_files = []
for cat in categories:
    cat_path=os.path.join(rootFolder,cat)
    if not os.path.isdir(cat_path):
        continue
    insts=[d for d in os.listdir(cat_path) if os.path.isdir(os.path.join(cat_path,d))]
    for inst in insts:
        inst_path=os.path.join(cat_path,inst)
        collections=[d for d in os.listdir(inst_path) if os.path.isdir(os.path.join(inst_path,d))]
        for coll in collections:
            subset = df[
                (df["collectionCode"] == coll) &
                 (df["assetCategory"] == cat) &
                 (df["format"] != ".csv")]
            if not subset.empty:
                meta_folder=os.path.join(inst_path,coll,"metadata")
                os.makedirs(meta_folder,exist_ok=True)
                subset_path=os.path.join(meta_folder,f"{coll}_metadata_{RUN_TIMESTAMP}.csv")
                subset.to_csv(subset_path,index=False)
                subset_files.append(subset_path)
                print(f"Collection metadata CSV generated: {subset_path}")
                # Add the subset CSV itself as a file in master df
                all_rows.append({
                    "fileName":os.path.basename(subset_path),
                    "documentId":f"{coll}_subset_{RUN_TIMESTAMP}",
                    "title":f"{coll}_subset_{RUN_TIMESTAMP}",
                    "institutionCode":inst,
                    "collectionCode":coll,
                    "institutionName":INSTITUTION_CODE_MAP.get(inst,inst),
                    "creator":meta["creator"],
                    "contributor":meta["contributor"],
                    "description":f"Subset metadata for {coll}",
                    "rightsHolder":meta["rightsHolder"],
                    "holdingInstitution":meta["holdingInstitution"],
                    "license":meta["license"],
                    "dateCreated":datetime.now().strftime("%Y:%m:%d %H:%M:%S"),
                    "format":".csv",
                    "subject":"metadata_subset",
                    "language":languageVar.get(),
                    "fullPath":subset_path,
                    "relativePath":os.path.relpath(subset_path,rootFolder),
                    "assetCategory":f"{cat}_metadata"
                })

df = pd.DataFrame(all_rows)  # refresh df to include subset CSVs

# ==================================================
# Write DAMSG output (master CSV/XLSX including subset CSVs)
# ==================================================
outRoot = os.path.join(rootFolder, "DAMSG_output", RUN_TIMESTAMP)
os.makedirs(outRoot, exist_ok=True)

# sanitize scan mode for filename: replace spaces, +, / with underscores, lowercase
scan_mode_str = scanModeVar.get().replace(" ", "_").replace("+", "plus").replace("/", "_").lower()

# append only the relevant code for single collection or single institution
if scanModeVar.get() == "Single Collection":
    extra_str = f"_{collectionVar.get().lower()}"
elif scanModeVar.get() == "All Collections (selected institution)":
    extra_str = f"_{institutionVar.get().lower()}"
else:
    extra_str = ""

# CSV output
if outputChoiceVar.get() in ("Both", "CSV only"):
    csv_path = os.path.join(outRoot,
                            f"digital_asset_inventory_{scan_mode_str}{extra_str}_{RUN_TIMESTAMP}.csv")
    df.to_csv(csv_path, index=False)

# Excel output
if outputChoiceVar.get() in ("Both", "Excel only"):
    xlsx_path = os.path.join(outRoot,
                             f"digital_asset_inventory_{scan_mode_str}{extra_str}_{RUN_TIMESTAMP}.xlsx")
    df.to_excel(xlsx_path, index=False)

print(f"Processing complete. DAMSG master export includes all files and newest subset CSVs.\nOutput folder: {outRoot}")

# ==================================================
# Open output file automatically
# ==================================================
def open_file(path):
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(path)
        elif system == "Darwin":
            subprocess.run(["open", path], check=True)
        else:  # Linux / others
            subprocess.run(["xdg-open", path], check=True)
    except Exception as e:
        print(f"Could not open {path}: {e}")

# CSV output
if outputChoiceVar.get() in ("Both", "CSV only"):
    csv_path = os.path.join(outRoot,
                            f"digital_asset_inventory_{scan_mode_str}{extra_str}_{RUN_TIMESTAMP}.csv")
    df.to_csv(csv_path, index=False)
    open_file(csv_path)  # <--- open CSV automatically

# Excel output
if outputChoiceVar.get() in ("Both", "Excel only"):
    xlsx_path = os.path.join(outRoot,
                             f"digital_asset_inventory_{scan_mode_str}{extra_str}_{RUN_TIMESTAMP}.xlsx")
    df.to_excel(xlsx_path, index=False)
    open_file(xlsx_path)  # <--- open Excel automatically