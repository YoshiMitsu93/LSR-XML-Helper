# LSR XML Helper

A dedicated XML editor for Los Santos RED configs.

It automatically detects all XML files in your chosen folder, loads them, and turns them into readable lists of categories and entries, so you don’t have to scroll through massive XML blocks manually.

## Key Features

- **Browse by entry type**  
  Vehicles, drugs, gangs, dealers, UI, gameplay settings, etc.  
  Select entries and edit any values directly: prices, durations, effects, spawn chances, strings, locations, and more.

- **Automatic backups**  
  Every time you save changes, the tool creates a timestamped XML backup in a `BackupXMLs` folder in the same location as the XMLs you are editing so you don’t lose work or overwrite something by mistake.

- **Built-in keyword search**  
  Search for anything (`"cocaine"`, `"mp5"`, `"Underground"`, etc.).  
  The tool scans all XML files and shows:
  - Which XML files contain the keyword  
  - Which entry types match  
  - Which specific entries contain it  
  - Which field(s) inside those entries matched  

  You can jump straight to the matching entry from the search results. The editor highlights the matching entry and the exact field that triggered the match, and you can edit values from there.

- **Duplicate entries**  
  Pick an existing entry, clone it, change the values, and the new version is added automatically under the same category.  
  You can also simply view entries without changing anything, or edit existing ones directly.

---

## LSR XML Helper – Important Information

### General

✔ View, edit, and duplicate entries inside XML files used by Los Santos RED.  
✔ You **must extract the ZIP** before using the tool. Do not run it directly from inside the archive.

---

## Windows “Protected your PC” Message

When you run the EXE for the first time, Windows may show:

> “Windows protected your PC”

This happens because the application is not commercially code-signed, not because it’s malicious.

To continue:

1. Click **“More info”**  
2. Click **“Run anyway”**

---

## Antivirus Warning Information

Some antivirus programs may show warnings or flag the file as suspicious, unknown, or potentially unsafe.

This is usually because:

- It is a new EXE with no reputation yet  
- It is unsigned  
- It edits files (which antivirus tools monitor closely)

If your antivirus blocks or deletes it:

1. Open your antivirus settings  
2. Add **`LSR-XML-Helper.exe`** to the allowed list / exclusions / safe list  

Once allowed, it should run normally.

---

## Auto Updates

- The tool automatically checks online for newer versions.  
- If an update is available, it downloads and replaces itself.  
- After the update, just launch the tool again as normal.

---

## How to Use

1. Launch **`LSR-XML-Helper.exe`**.  
2. Select the folder containing your Los Santos RED XML files.  
3. Pick an XML file to view or edit.  
4. Make your changes and save.

The selected folder is remembered automatically for next time.

---

## What This Tool Does *Not* Do

This tool does **not**:

- Modify anything outside your chosen XML folder  
- Install files onto your system  
- Change registry settings  
- Require admin permissions  
- Send data anywhere  
- Stay running after exit  

It only edits XML files that **you explicitly choose** in the folder you point it at.
