# MLS Jobs Automation (Advanced)

This Google Apps Script automates the creation and management of job folders and notes in Google Drive based on rows in a Google Sheet. It extends the original trial by adding subfolders, a config sheet, and improved notes formatting.

## Features
- **Automatic folder creation** under `MLS Jobs/2025/<Bid#> <Client>`
- **Subfolders**: Creates `Field`, `Drafting`, and `Deliverables` inside each job folder
- **Job Notes file**: Generates/updates a `Job Notes.txt` file with row details and timestamp
- **No duplicate folders**: Uses stored Folder ID to prevent re-creation
- **Config sheet support**: Root folder ID is stored in the `Config` tab of the Sheet instead of being hardcoded
- **Writeback**: Saves both the main folder URL and the Deliverables subfolder URL back to the Sheet
- **Custom menu**: `MLS Jobs` menu with options to process all rows or the current row

## Setup Instructions
1. Create a Google Sheet with:
   - `Info` tab (headers: Bid# | Client | TMK | Address | Status)
   - `Config` tab with:
     - Cell A2: `ROOT_FOLDER_ID`
     - Cell B2: `<Paste your Google Drive folder ID here>`
2. Open `Extensions → Apps Script`
3. Copy the code from `MLS_Jobs_Automation.gs` into the script editor
4. Save and install an **onEdit trigger**:
   - Function: `onEdit`
   - Event type: From spreadsheet → On edit
5. Add or update rows in the `Info` tab and the script will create/update job folders and notes.

## Handling Duplicates & Deletions

**Duplicate Folders**
The script prevents duplicates by storing the Folder ID in the Info tab (column I by default).
If a row already has a stored folder ID, the script reuses that folder instead of creating a new one.
If the folder name changes (e.g., client name updated), the script renames the existing folder instead of duplicating it.

**Duplicate Job Notes Files** 
Inside each folder, before creating a new Job Notes.txt, the script looks for any existing file with the same name.
If found, the old file is moved to trash.
A new, updated Job Notes.txt is then created.
This ensures there’s always exactly one notes file per job folder.

**Deletion Safety**
Only job notes inside the job’s folder are deleted/overwritten.
No unrelated files or folders in Google Drive are touched.
