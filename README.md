# Loot Master App

**Two windows for quick D\&D loot management**

---

## Table of Contents

1. [About](#about)
2. [Current Functionality](#current-functionality)
3. [Installation](#installation)
4. [Usage](#usage)
5. [Configuration & Data Format](#configuration--data-format)
6. [Development Roadmap](#development-roadmap)
7. [Future Plans](#future-plans)
8. [Contributing](#contributing)
9. [License](#license)

---

## About

The **Loot Master App** is a lightweight PySide6 desktop application designed for D\&D game masters to:

* **Generate randomized loot** from predefined loot boxes
* **Assign items** directly into player inventories
* **View and manage** party or individual inventories
* **Manually add, trade, or drop items** in player inventories
* **Persist data** in an Excel spreadsheet (`ErwinLootTable.xlsx`), with robust error handling and options for auto-update or manual read/write

It provides two synchronized windows:

1. **Loot Box Generator** – select a loot box, roll for loot, take or drop items
2. **Player Inventory** – view aggregated inventories by player or entire party, with manual add, trade, and drop actions

---

## Current Functionality

* **Real Excel I/O**: Loads `Loot`, `Loot box sizes`, and `Players` sheets. Optionally writes back updated player inventories.
* **Roll Logic**: Filters items by scarcity and value ranges, then randomly samples based on box definitions.
* **Dynamic UI**: Two windows stay in sync. Taking loot immediately updates the inventory view.
* **Auto-update Toggle & Excel Options**: Use the Excel Options dialog to enable auto-update, or manually read/write Excel data on demand.
* **Manual Add, Trade, Drop**: Add items, trade between players, or drop items from inventory with intuitive popups and sliders.
* **Error Handling**: If the Excel file is missing, it is auto-created. User-friendly dialogs for file errors and invalid actions.
* **One-decimal Precision**: All weights and values are rounded to one decimal place.
* **Clean, Styled Tables**: Tables with alternating row colors, rounded action buttons, and clear totals.
* **No Row Index**: Row indices are hidden in all tables and not saved to Excel.

---

## Installation

1. Clone or download this repository.
2. Ensure you have Python 3.8+ installed (64-bit recommended).
3. Install dependencies:

   ```bash
   pip install PySide6 pandas openpyxl
   ```
4. Place your `ErwinLootTable.xlsx` next to `loot_master_app.py`.

---

## Usage

```bash
python loot_master_app.py
```

* **Loot Box Generator Window**

  1. (Optional) Open **Excel Options** to enable **Auto-update Excel** or use manual **Read/Write**.
  2. Select a **Loot Box** from the drop-down.
  3. Click **Roll** to populate the loot table.
  4. (Optional) Select **Player** and click **Take** on any row, or **Take All**.
  5. Click **Drop** to remove items from the current roll or from player inventory (with quantity slider).

* **Player Inventory Window**

  1. Select a **Player** or **Party** to view aggregated inventory.
  2. Use the **+** button to manually add items to a player's inventory (popup with item/quantity selection).
  3. Use the **Trade** button to move items between players (popup with slider).
  4. Use the **Drop** button to remove items from inventory (popup with slider).
  5. Totals for weight and value update automatically.
  6. Actions are disabled when "Party" is selected.

* **Excel Options Dialog**

  - Access via the menu or settings button.
  - Toggle **Auto-update** to write changes to Excel in real time.
  - Use **Read** to reload all data from Excel, or **Write** to save all inventories on demand (when auto-update is off).

* **Error Handling**

  - If the Excel file is missing, it is auto-created with default sheets.
  - User-friendly dialogs appear for file errors, permission issues, or invalid actions.

Closing either window via the red **X** will exit the entire application.

---

## Configuration & Data Format

### Excel Schema (`ErwinLootTable.xlsx`)

* **Loot** sheet:

  * `Item` (string)
  * `Description` (string)
  * `Value(GP)` (float)
  * `Max` (int)
  * `Weight` (float)
  * `Item scarecity` (int)

* **Loot box sizes** sheet:

  * `Loot box name` (string)
  * `Max total items` (int)
  * `Min box value` (float)
  * `Max box value` (float)
  * `Min scarecity` (int)
  * `Max scarecity` (int)

* **Players** sheet (two header rows):

  * Top row: player names
  * Second row: columns `Loot` and `Qty`
  * Rows: current items and quantities per player

---

## Development Roadmap

* **Modularize code** into separate modules (`data.py`, `ui.py`, `logic.py`).
* **Unit Tests** for roll logic and Excel I/O.
* **Configuration file** (`config.yaml`) for customizing themes and default paths.

---

## Future Plans

1. **Editable Descriptions**: Hover tooltips showing item descriptions loaded from Excel.
2. **Save/Load Profiles**: Support multiple campaign files beyond one Excel.
3. **Trade & Drop**: Implement action buttons in the inventory window.
4. **Export Reports**: Generate PDF or CSV summaries of inventory states.
5. **Plugin System**: Allow custom roll formulas or additional loot sources.
6. **Cross-platform Packaging**: Create single-file executables for Windows, macOS, and Linux.

---

## Contributing

Contributions are welcome! Please:

1. Fork the repository.
2. Create a new branch: `git checkout -b feature/YourFeature`.
3. Commit your changes with clear messages.
4. Submit a pull request detailing your additions.

---

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
