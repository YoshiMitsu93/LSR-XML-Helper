# LSR XML Helper

A dedicated XML editor for Los Santos RED configs.

It automatically detects all XML files in your chosen folder, loads them, and turns them into readable lists of categories and entries, so you don’t have to scroll through massive XML blocks manually.

## Key Features

- **Browse by entry type**  
  Vehicles, drugs, gangs, dealers, UI, gameplay settings, etc.  
  Select entries and edit any values directly: prices, durations, effects, spawn chances, strings, locations, and more.

- **Automatic backups**  
  Every time you save changes, the tool creates a timestamped XML backup inside a `BackupXMLs` folder, located next to the XML files you are editing.

- **Built-in keyword search**  
  Search for values like `"cocaine"`, `"mp5"`, `"Underground"`, etc.  
  The tool scans all XML files and shows:
  - Which XML files contain the keyword  
  - Matching entry types  
  - Matching entries  
  - The exact fields where the match was found  

  From the results screen, you can jump directly into editing that entry.

- **Duplicate or modify existing entries**  
  Select any entry, clone it, edit its values, and the new version will be inserted into the XML under the correct category.

- **Edit history**  
  Every change you make, new entries or modified fields is saved into a JSON-based change log.
  This means you can:
  - Review all saved changes at any time  
  - Apply them again later  
  - Delete individual changes  
  - Delete all saved changes for a file  
  - Re-apply your modifications if XMLs are replaced or updated  

  This is ideal when updating to newer LSR releases:  
  simply load fresh XML files and re-apply all your saved edits automatically.

---

## LSR XML Helper – Important Information

### General

✔ View, edit, and duplicate entries inside XML files used by Los Santos RED.  
✔ You **must extract the ZIP** before using the tool.

---

## Windows “Protected your PC” Message

Windows may show:

> “Windows protected your PC”

This happens because the file is unsigned, not because it is harmful.

To continue:

1. Click **“More info”**  
2. Click **“Run anyway”**

---

## Antivirus Warning Information

Some antivirus programs may flag the file as unrecognized or suspicious because:

- It is new and unsigned  
- It edits local files  
- It downloads updates  

If blocked:

1. Open antivirus settings  
2. Add **`LSR-XML-Helper.exe`** to your allowed list  

---

## Auto Updates

- The tool automatically checks online for newer versions.  
- If an update is found, it downloads and replaces itself.  
- After updating, simply launch it again.

---

## How to Use

1. Launch **`LSR-XML-Helper.exe`**  
2. Select your Los Santos RED XML folder  
3. View, add, or edit entries  
4. Save when finished  
5. Saved edits remain stored for future use

Your selected folder is remembered automatically.

---

## What This Tool Does *Not* Do

The tool does **not**:

- Modify files outside your chosen folder  
- Install anything  
- Change registry settings  
- Require admin permissions  
- Upload or share any data  
- Run after you close it  

It only reads/writes XML files that **you explicitly select**.
