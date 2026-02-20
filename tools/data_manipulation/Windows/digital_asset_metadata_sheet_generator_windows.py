#!/usr/bin/env python3
import os
import subprocess
import pandas as pd
from datetime import datetime
from tkinter import (
    Tk, Label, Button, StringVar, OptionMenu,
    filedialog, messagebox, DISABLED, NORMAL
)
from PIL import Image, ExifTags
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
outputChoices = ["Both","Excel only","CSV only"]
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
def selectRootFolder():
    p = filedialog.askdirectory()
    if p:
        rootFolderVar.set(p)
        rootLabel.config(text=p)

def selectMappingCSV():
    global mappingDF
    p = filedialog.askopenfilename(filetypes=[("CSV files","*.csv")])
    if not p:
        return
    try:
        df = pd.read_csv(p)
    except Exception as e:
        messagebox.showerror("CSV Error",str(e))
        return
    required = {"institutionCode","collectionCode","creator","contributor","license","rightsHolder","holdingInstitution","description"}
    missing = required - set(df.columns)
    if missing:
        messagebox.showerror("Mapping CSV Error",f"Missing required columns:\n{', '.join(sorted(missing))}")
        return
    mappingDF = df
    mappingFileVar.set(p)
    mappingLabel.config(text=p)
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
Button(root,text="Select Metadata Mapping CSV",command=selectMappingCSV,bg="orange").pack(fill="x", padx=20, pady=5)
mappingLabel=Label(root,text="",wraplength=800,anchor="w")
mappingLabel.pack()

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
    for r,_,files in os.walk(collectionRoot):
        for f in files:
            if f.lower().endswith(extensions):
                full=os.path.join(r,f)
                rel=os.path.relpath(full,rootFolder)
                base=os.path.splitext(f)[0]
                fmt=os.path.splitext(f)[1].lower()
                subject = "metadata" if fmt==".csv" else ""
                rows.append({
                    "documentId":f"{base}_{collectionCode}_{hash(rel)&0xffff}",
                    "title":f"{base} ({collectionCode})",
                    "institutionCode":institutionCode,
                    "collectionCode":collectionCode,
                    "institutionName":INSTITUTION_CODE_MAP.get(institutionCode,institutionCode),
                    "creator":meta["creator"],
                    "contributor":meta["contributor"],
                    "description":f'{meta["description"]} ({collectionCode})',
                    "rightsHolder":meta["rightsHolder"],
                    "holdingInstitution":meta["holdingInstitution"],
                    "license":meta["license"],
                    "dateCreated":getDateCreated(full),
                    "format":fmt,
                    "subject":subject,
                    "language":languageVar.get(),
                    "fileName":f,
                    "fullPath":full,
                    "relativePath":rel,
                    "assetCategory":categoryRoot
                })
    return rows

categories = [d for d in os.listdir(rootFolder) if os.path.isdir(os.path.join(rootFolder,d))]

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
            subset=df[(df["collectionCode"]==coll)&(df["format"]!=".csv")]
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
                    "assetCategory":"collection_metadata"
                })

df = pd.DataFrame(all_rows)  # refresh df to include subset CSVs

# ==================================================
# Write DAMSG output (master CSV/XLSX including subset CSVs)
# ==================================================
outRoot=os.path.join(rootFolder,"DAMSG_output",RUN_TIMESTAMP)
os.makedirs(outRoot,exist_ok=True)

if outputChoiceVar.get() in ("Both","CSV only"):
    df.to_csv(os.path.join(outRoot,f"imageFilesList_{RUN_TIMESTAMP}.csv"),index=False)
if outputChoiceVar.get() in ("Both","Excel only"):
    xlsx=os.path.join(outRoot,f"imageFilesList_{RUN_TIMESTAMP}.xlsx")
    df.to_excel(xlsx,index=False)
    system=platform.system()
    if system=="Windows":
        os.startfile(xlsx)
    elif system=="Darwin":
        subprocess.run(["open",xlsx])
    else:
        subprocess.run(["xdg-open",xlsx])

print(f"Processing complete. DAMSG master export includes all files and newest subset CSVs.\nOutput folder: {outRoot}")