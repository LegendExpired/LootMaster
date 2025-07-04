# Error Handling Plan for Loot Master App

This document outlines error situations that still need to be handled in the Loot Master App, organized by category. For each, handling strategies are suggested.

## 1. OS and I/O Related Errors
- **Excel file is open/locked by another program**
  - Show a user-friendly error dialog explaining the file is in use and suggest closing it in Excel or other programs.
  - Retry or allow the user to cancel the operation.
- **Permission denied when reading/writing Excel file**
  - Show an error dialog with the file path and suggest running the app with appropriate permissions or choosing a different location.
- **File not found in directory**
  - Show a popup saying the file was not found (printing out the directory the file is expected in as part of the popup) and then asking the user if they would like to create the file.
- **Disk full or out of space**
  - Notify the user that saving failed due to insufficient disk space.
- **File path too long or invalid characters**
  - Validate file paths before use and show a clear error if invalid.

## 2. Excel File Format Related Errors
- **Missing required sheets ("Loot", "Loot box sizes", "Players")**
  - Detect missing sheets and offer to auto-create them with default data, or prompt the user to fix the file.
- **Corrupted or unreadable Excel file**
  - Catch exceptions from pandas/openpyxl, show an error dialog, and offer to restore from backup or create a new file.
- **Unexpected column names or missing columns**
  - Validate columns on load; if missing or renamed, show a dialog listing the problem and offer to auto-fix or abort.
- **MultiIndex header issues in Players sheet**
  - Detect header format problems and prompt the user to fix or delete and recreate the sheet.
- **Non-numeric or invalid data in numeric columns (Qty, Value, Weight, Scarcity, etc.)**
  - Validate data types on load; highlight or skip invalid rows, and inform the user.
- **Inventory items that do not exist in the loot database**
  - Ensure all inventory items exist in the loot database.  If an item is found in any players inventory but not in the loot database, provide a pop up for the user telling them which item(s) (if there is more than 1 even) are missing from the loot table, and ask them to insert the items (case sensitive) in that table.  The popup should give an option to retry the load or to exit.  If retry it should recheck if the problem is resolved.

## 3. User Input Related Errors
- **Entering negative or zero quantities**
  - Prevent in UI; validate before saving to inventory or Excel.
- **Adding items not present in the loot table**
  - Restrict selection to valid items; if not, show an error and prevent the action.
- **Attempting to drop/trade/add items when "Party" is selected**
  - Disable or block these actions in the UI and show an informative message.
- **Exceeding maximum allowed quantity for an item**
  - Enforce max limits in the UI and validate before saving.
- **Rapid repeated actions (e.g., double-clicking buttons)**
  - Debounce actions or disable buttons during processing to prevent duplicate entries.

## 4. Application Logic/Sync Errors
- **Inventory and Excel file out of sync**
  - Warn the user if the in-memory data differs from the file and offer to reload or overwrite.
- **Concurrent edits from multiple app instances**
  - Detect file changes on disk and prompt the user to reload or merge changes.

## 5. Miscellaneous
- **Unhandled exceptions**
  - Add a global exception handler to show a user-friendly error dialog and log details for debugging.
- **Missing dependencies (pandas, openpyxl, PySide6)**
  - On startup, check for required packages and show a clear message if missing, with install instructions.
