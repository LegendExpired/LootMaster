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
* **Persist data** in an Excel spreadsheet (`ErwinLootTable.xlsx`)

It provides two synchronized windows:

1. **Loot Box Generator** – select a loot box, roll for loot, take or drop items
2. **Player Inventory** – view aggregated inventories by player or entire party

---

## Current Functionality

* **Real Excel I/O**: Loads `Loot`, `Loot box sizes`, and `Players` sheets. Optionally writes back updated player inventories.
* **Roll Logic**: Filters items by scarcity and value ranges, then randomly samples based on box definitions.
* **Dynamic UI**: Two windows stay in sync. Taking loot immediately updates the inventory view.
* **Auto-update Toggle**: Check a box to write inventory changes back to the Excel file in real time.
* **One-decimal Precision**: All weights and values are rounded to one decimal place.
* **Clean, Styled Tables**: Tables with alternating row colors, rounded action buttons, and clear totals.

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

  1. (Optional) Check **Auto-update Excel** to persist immediately.
  2. Select a **Loot Box** from the drop-down.
  3. Click **Roll** to populate the loot table.
  4. (Optional) Select **Player** and click **Take** on any row, or **Take All**.
  5. Dropped items are removed from the current roll only.

* **Player Inventory Window**

  1. Select a **Player** or **Party** to view aggregated inventory.
  2. Totals for weight and value update automatically.
  3. (Future) Buttons allow trading or dropping items.

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
