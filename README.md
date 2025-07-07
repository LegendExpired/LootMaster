# Loot Master App

> **Note:** The code readability needs massive work. As this was a rushed and quick project, someone can refactor the code later for readability and maintainability.

## Version

Current version: **1.0.0**

See [CHANGELOG.md](CHANGELOG.md) for a full list of changes.

## Overview
Loot Master is a D&D loot and inventory management app with a PySide6 GUI and Excel integration. It allows you to generate loot boxes, manage player inventories, and persist all data in an Excel file.

<img src="resources/loot_box_icon.png" alt="Loot Master Icon" width="64" height="64" />

## Features
- Two windows: Loot Box Generator and Player Inventory
- Real-time updates to inventory and loot
- Auto-update Excel file (loot_table.xlsx) with all changes
- Robust error handling for Excel file and data issues
- User-friendly dialogs for all errors and user actions

## How to Use the Application

### 1. Running the App
- Make sure you have Python 3.8+ installed.
- Install dependencies (see below).
- Run the app:
  ```sh
  python dnd_master_loot_gen.py
  ```
- The app will create a default `loot_table.xlsx` if it does not exist.

### 2. Using the GUI
- **Loot Box Generator Window:**
  - Select a loot box type and player, then click 'Roll' to generate loot.
  - Use 'Take' or 'Take All' to add loot to the selected player's inventory.
  - 'Excel Options' lets you enable/disable auto-update and manually read/write Excel.
- **Player Inventory Window:**
  - Select a player or 'Party' to view inventory.
  - Use '+' to add items to a player's inventory.
  - Use 'Trade' or 'Drop' to move or remove items (cannot drop/trade as 'Party').

### 3. Editing the Excel File
- The app uses `loot_table.xlsx` in the same folder as the script.
- You can edit the Excel file directly, but keep the sheet names and columns as generated.
- If you add new items or players, ensure all required columns are present and numeric values are valid.
- If the file is corrupted or missing columns/sheets, the app will prompt you to fix or regenerate it.

## App Icon
- Please use `loot_box_icon.ico` as the app icon. Place this file in the same directory as the script.

## How to Build the Application

1. **Clone or Download the Repository**
2. **Install Python 3.8+**
3. **Install Dependencies:**
   ```sh
   pip install -r requirements.txt
   # or, if requirements.txt is missing:
   pip install PySide6 pandas openpyxl
   ```
4. **Run the App:**
   ```sh
   python dnd_master_loot_gen.py
   ```

### Build a Standalone Executable with PyInstaller

1. **Install PyInstaller:**
   ```sh
   pip install pyinstaller
   ```
2. **Build the App as a Single Executable:**
   ```sh
   pyinstaller --onefile --windowed --icon resources/loot_box_icon.ico --add-data "resources/loot_box_icon.ico;resources" dnd_master_loot_gen.py
   ```
   - The `--onefile` flag creates a single executable.
   - The `--windowed` flag prevents a console window from appearing (for GUI apps).
   - The `--icon` flag sets the app icon (ensure `loot_box_icon.ico` is present).
3. **Find the Executable:**
   - The output will be in the `dist/` folder as `dnd_master_loot_gen.exe` (Windows) or the appropriate binary for your OS.

## Developer Environment Setup

1. **Recommended Tools:**
   - Visual Studio Code or PyCharm
   - Python 3.8+
   - Git
2. **Install Dev Dependencies:**
   - (Optional) Use a virtual environment:
     ```sh
     python -m venv venv
     source venv/bin/activate  # On Windows: venv\Scripts\activate
     ```
   - Install packages:
     ```sh
     pip install -r requirements.txt
     # or
     pip install PySide6 pandas openpyxl
     ```
3. **Linting and Formatting:**
   - Use `flake8` or `pylint` for linting.
   - Use `black` for code formatting.
4. **Testing:**
   - Manual testing via the GUI is recommended.
   - Add unit tests for new features if possible.

## Troubleshooting
- If you see errors about missing dependencies, install them with pip.
- If the Excel file is locked, close it in Excel and retry.
- If you get format errors, use the app's dialogs to regenerate or fix the file.

---

For further improvements, see the `errorhandling.md` for a list of known error cases and strategies.

## License

MIT License

Copyright (c) 2025 LegendExpired

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the “Software”), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell      
copies of the Software, and to permit persons to whom the Software is          
furnished to do so, subject to the following conditions:                       

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.                                

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR     
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,       
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE    
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER         
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,  
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE  
SOFTWARE.
