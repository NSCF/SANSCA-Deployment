#!/usr/bin/env python3
import os
import subprocess
import pandas as pd
from datetime import datetime
from tkinter import (
    Tk, Canvas, Label, Button, StringVar, OptionMenu,
    filedialog, messagebox, DISABLED, NORMAL
)
from PIL import Image, ExifTags, ImageTk
import platform
import sys
import pathlib
import hashlib

# ==================================================
# Run timestamp
# ==================================================
RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

# ==================================================
# Institution display map (optional)
# ==================================================
INSTITUTION_CODE_MAP = {
    "ISAM": "Iziko Museum of South Africa",
    "DNMNH": "Ditsong National Museum of Natural History"
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
# Deterministic documentId generator for image/assets
# ==================================================
def generate_document_id(institution_code,collection_code, base_name, relative_path, length=8):
    relative_path = relative_path.replace("\\", "/")
    clean_inst = institution_code.replace("_", "")
    clean_collection = collection_code.replace("_", "")
    clean_base = base_name.replace("_", "")
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
root.geometry("850x800")

Label(root,text="Select Root Folder and Mapping CSV",font=("Arial",12,"bold")).pack(pady=10)

rootFolderVar = StringVar()
mappingFileVar = StringVar()
institutionVar = StringVar()
collectionVar = StringVar()
fileFilterVar = StringVar()
outputChoiceVar = StringVar()
scanModeVar = StringVar()

fileFilters = ["All","TIFF only","RAW only","JPEG only"]
outputChoices = ["CSV only","Excel only","Both"]
scanModes = ["Single Collection","All Collections (selected institution)","All Institutions + Collections"]

fileFilterVar.set(fileFilters[0])
outputChoiceVar.set(outputChoices[0])
scanModeVar.set(scanModes[0])

mappingDF = None

# ==================================================
# UI functions
# ==================================================
script_dir = pathlib.Path(__file__).parent.resolve()
logo_path = script_dir / "nscf_logo_crop.png"

canvas = Canvas(root, width=300, height=150, bg="white", highlightthickness=0)
canvas.pack(pady=10)

try:
    logo_img = Image.open(logo_path).convert("RGB")
    logo_img = logo_img.resize((300, 150), Image.Resampling.LANCZOS)
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

    global mappingDF
    mapping_path = os.path.join(p, "DAMSG_mapping", "damsg_tool_collection_mapping_master_2026.csv")
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

    required = {"institutionCode","collectionCode","creator","contributor",
                "license","rightsHolder","holdingInstitution","description"}
    missing = required - set(df.columns)
    if missing:
        messagebox.showerror("Mapping CSV Error",
                             f"Missing required columns:\n{', '.join(sorted(missing))}")
        mappingDF = None
        return

    mappingDF = df
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
# File scanning logic
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

                asset_category = categoryRoot
                if os.path.sep + "metadata" + os.path.sep in full:
                    asset_category = f"{categoryRoot}_metadata"

                parts = base.split("_")
                view_code = parts[1] if len(parts) > 1 else ""
                view_desc = VIEW_CODE_MAP.get(view_code.lower(), "")
                description_text = view_desc if view_desc else meta.get("description", collectionCode) 

                if fmt == ".csv":
                    doc_id = generate_metadata_document_id(institutionCode,collectionCode, categoryRoot, rel)
                else:
                    doc_id = generate_document_id(institutionCode, collectionCode, base, rel)

                row_data = {
                    "documentId": doc_id,
                    "title": base,
                    "institutionCode": institutionCode,
                    "collectionCode": collectionCode,
                    "institutionName": INSTITUTION_CODE_MAP.get(institutionCode, institutionCode),
                    "description": description_text,
                    "dateCreated": getDateCreated(full),
                    "format": fmt,
                    "additionalNames": "",  # start empty
                    "subject": "Metadata" if fmt == ".csv" else meta.get("subject", ""),
                    "fileName": f,
                    "fullPath": full,
                    "relativePath": rel,
                    "assetCategory": asset_category,
                    "scanModeApplied": scanMode
                }

                # Add ALL mapping columns automatically
                for col in mappingDF.columns:
                    if col not in row_data or (col == "additionalNames" and not row_data["additionalNames"].strip()):
                        row_data[col] = meta.get(col, "")

                rows.append(row_data)   

    return rows

# ==================================================
# Master file paths
# ==================================================
master_csv = os.path.join(rootFolder, "DAMSG_output", "digital_asset_inventory_master.csv")
master_xlsx = os.path.join(rootFolder, "DAMSG_output", "digital_asset_inventory_master.xlsx")
os.makedirs(os.path.dirname(master_csv), exist_ok=True)

if os.path.isfile(master_csv):
    master_df = pd.read_csv(master_csv)
else:
    master_df = pd.DataFrame()

# ==================================================
# Scan, generate subset CSVs, and append newest metadata
# ==================================================
categories = [d for d in os.listdir(rootFolder) if os.path.isdir(os.path.join(rootFolder, d))]

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

            try:
                meta = mappingDF[(mappingDF["institutionCode"] == inst) & (mappingDF["collectionCode"] == coll)].iloc[0]
            except IndexError:
                continue

            all_rows.extend(scan_collection(cat, inst, coll, meta))

            subset_rows = [
                r for r in all_rows
                if r["collectionCode"] == coll and r["assetCategory"] == cat and r["format"] != ".csv"
            ]
            if subset_rows:
                meta_folder = os.path.join(inst_path, coll, "metadata")
                os.makedirs(meta_folder, exist_ok=True)
                subset_path = os.path.join(meta_folder, f"{coll}_metadata_{RUN_TIMESTAMP}.csv")
                pd.DataFrame(subset_rows).to_csv(subset_path, index=False)
                print(f"Collection metadata CSV generated: {subset_path}")

                row_data = {
                    "fileName": os.path.basename(subset_path),
                    "documentId": generate_metadata_document_id(inst, coll, cat, os.path.relpath(subset_path, rootFolder)),
                    "title": os.path.basename(subset_path),
                    "institutionCode": inst,
                    "collectionCode": coll,
                    "institutionName": INSTITUTION_CODE_MAP.get(inst, inst),
                    "creator": meta["creator"],
                    "contributor": meta["contributor"],
                    "description": meta.get("description", ""),
                    "rightsHolder": meta["rightsHolder"],
                    "holdingInstitution": meta["holdingInstitution"],
                    "license": meta["license"],
                    "dateCreated": datetime.now().strftime("%Y:%m:%d %H:%M:%S"),
                    "format": ".csv",
                    "subject": "Metadata",
                    "fullPath": subset_path,
                    "relativePath": os.path.relpath(subset_path, rootFolder),
                    "assetCategory": f"{cat}_metadata",
                    "scanModeApplied": scanMode,
                    "additionalNames": ""
                }

                if "additionalNames" in mappingDF.columns and not row_data["additionalNames"].strip():
                    row_data["additionalNames"] = meta.get("additionalNames", "")

                all_rows.append(row_data)


# ==================================================
# Convert all_rows to DataFrame safely
# ==================================================

expected_columns = [
    "documentId","title","fileName","relativePath","fullPath",
    "format","assetCategory","dateCreated","scanModeApplied",
    "institutionCode","collectionCode","institutionName",
    "description","additionalNames","creator","contributor",
    "license","rightsHolder","holdingInstitution","subject"
]

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
        ~new_rows_df["documentId"].isin(master_df["documentId"])
    ]

# Preserve description for CSV metadata safely
required_cols = {"format", "description", "documentId", "institutionCode", "collectionCode"}
if not new_rows_df.empty and required_cols.issubset(new_rows_df.columns):
    mask_csv = new_rows_df["format"] == ".csv"

    if not master_df.empty and "documentId" in master_df.columns:
        existing_desc = master_df.set_index("documentId")["description"].to_dict()

        def preserve_or_generate_desc(row):
            doc_id = row["documentId"]
            if doc_id in existing_desc and existing_desc[doc_id].strip():
                return existing_desc[doc_id]
            mode = row.get("scanModeApplied", "")
            if mode == "Single Collection":
                return f"Metadata file for {row['institutionCode']}_{row['collectionCode']}"
            elif mode == "All Collections (selected institution)":
                return f"Metadata file for {row['institutionCode']}"
            else:
                return "Metadata file for all collections"

        new_rows_df.loc[mask_csv, "description"] = new_rows_df.loc[mask_csv].apply(preserve_or_generate_desc, axis=1)
    else:
        def generate_desc(row):
            mode = row.get("scanModeApplied", "")
            if mode == "Single Collection":
                return f"Metadata file for {row['institutionCode']}_{row['collectionCode']}"
            elif mode == "All Collections (selected institution)":
                return f"Metadata file for {row['institutionCode']}"
            else:
                return "Metadata file for all collections"

        new_rows_df.loc[mask_csv, "description"] = new_rows_df.loc[mask_csv].apply(generate_desc, axis=1)

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
    "documentId","title","fileName","relativePath","fullPath",
    "format","assetCategory","dateCreated","scanModeApplied"] if col in updated_master_df.columns]

ordered_columns = system_columns + [col for col in mapping_columns if col not in system_columns]
updated_master_df = updated_master_df[ordered_columns]

# ==================================================
# Write master CSV/Excel
# ==================================================
updated_master_df.to_csv(master_csv, index=False)
updated_master_df.to_excel(master_xlsx, index=False)
print(f"Processing complete. Master inventory updated: {master_csv}")

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

if outputChoiceVar.get() in ("Both","CSV only"):
    open_file(master_csv)
if outputChoiceVar.get() in ("Both","Excel only"):
    open_file(master_xlsx)