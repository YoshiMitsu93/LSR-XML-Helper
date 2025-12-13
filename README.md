# LSR XML Helper

A dedicated XML editor for Los Santos RED configs.

It automatically detects all XML files in your chosen folder, loads them, and turns them into readable lists of categories and entries, so you don’t have to scroll through massive XML blocks manually.

## Key Features

- **Browse by entry type**  
  Vehicles, drugs, gangs, dealers, UI, gameplay settings, etc.  
  Select entries and edit any values directly: prices, durations, effects, spawn chances, strings, locations, and more.

- **Automatic backups**  
  Every time you save changes, the tool creates a timestamped XML backup inside a `BackupXMLs` folder, located under an `LSR-XML-Helper` helper folder next to the XML files you are editing.

- **Backup restore**  
  A dedicated **Restore Backups** menu lets you:
  - Browse backups for a single XML file  
  - View **all backups across all XML files** in one combined list  
  - Restore any backup with safety checks  
  - Automatically create a fresh backup before restoring (so you can undo a restore if needed)

- **Built-in keyword search**  
  Search for values like `"cocaine"`, `"mp5"`, `"Underground"`, etc.  
  The tool scans all XML files and shows:
  - Which XML files contain the keyword  
  - Matching entry types  
  - Matching entries  
  - The exact fields where the match was found  

  The search supports multiple words and quoted phrases in the same query.  
  From the results screen, you can jump directly into editing that entry, with matching fields highlighted.

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

- **Saved-edits summaries (exportable text)**  
  You can export a clear, readable **text summary** of your saved edits.
  - Export a summary for **all XML files** with saved edits  
  - Or export a summary for **a single XML file**  
  - Includes `Pending` / `Committed` status grouping and key change details  

  This is useful for:
  - Quickly reviewing what has been changed without opening XMLs  
  - Sharing a clean changelog with server owners or collaborators  
  - Keeping a record of edits before updates, restores, or resets

- **Shared config packs (export & import)**  
  - Export your saved edits into a single JSON “config pack”, covering one or many XML files  
  - Other users can import the pack and have the edits added as **pending changes**  
  - Imported edits are **merged**, not overwritten, preserving existing custom edits  
  - Removes the need to manually copy and paste XML sections

- **Review saved edits menu**  
  A dedicated menu lets you:
  - View all saved edits for each XML file, with `Pending` / `Committed` status tags  
  - Apply all changes at once (with automatic backups)  
  - Apply or delete individual changes  
  - Clear all saved changes for a file  
  - Run a dry-run test that applies changes in memory only  
  - Save the XML, create a backup, and mark changes as committed

- **Settings & Info screen**  
  A separate screen where you can:
  - View the current tool version, root XML folder, AppData config folder, and helper root  
  - Open the main XML folder, `XML-Edits`, `BackupXMLs`, and `Shared-Configs` folders instantly  
  - Toggle **automatic update checks**  
  - Toggle **auto-use last XML folder**

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

- The tool can check online for newer versions  
- If an update is found, it can download and replace itself  
- After updating, simply launch it again  
- Automatic update checks can be turned **ON/OFF** from the **Settings & Info** screen

---

## How to Use

1. Launch **`LSR-XML-Helper.exe`**  
2. Select your Los Santos RED XML folder  
3. View, add, or edit entries  
4. Use **Review saved edits**, **Saved-edits summaries**, and **Shared config packs** if you want history or shareable changes  
5. Save when finished  
6. Saved edits remain stored for future use

Your selected folder is remembered automatically (and can be auto-reused if you enable **auto-use last XML folder** in Settings & Info).

---

## What This Tool Does *Not* Do

The tool does **not**:
- Modify files outside your chosen folder  
- Install anything  
- Change registry settings  
- Require admin permissions  
- Upload or share any data  
- Run after you close it  

It only reads and writes XML files that **you explicitly select**.
