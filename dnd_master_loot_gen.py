#!/usr/bin/env python3
"""
Minimal PySide6 App with Two Windows

This application opens two separate windows:
1. Loot Box Generator
2. Player Inventory

It also loads Excel data into three pandas tables:
- items_df: master loot items
- boxes_df: loot box definitions
- inventory_df: flattened player inventories

Run:
    python minimal_pyside_app.py

Requires:
    PySide6, pandas, openpyxl
"""

import sys
import os
import pandas as pd
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel, QWidget, QVBoxLayout
from PySide6.QtCore import Qt

# Path to Excel file
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "ErwinLootTable.xlsx")


def load_data(filepath):
    """
    Load loot data from Excel into three DataFrames:
      - items_df: columns [Item, Description, Value, MaxQty, Weight, Scarcity]
      - boxes_df: columns [BoxName, MaxItems, MinValue, MaxValue, MinScarcity, MaxScarcity]
      - inventory_df: flattened inventory records [Player, Item, Qty, Value, Weight, Scarcity]
    """
    # Items sheet
    items_df = pd.read_excel(filepath, sheet_name="Loot").dropna(subset=["Item"])
    items_df.rename(
        columns={"Value(GP)": "Value", "Max": "MaxQty", "Item scarecity": "Scarcity"},
        inplace=True,
    )

    # Loot box definitions sheet
    boxes_df = pd.read_excel(filepath, sheet_name="Loot box sizes")
    boxes_df.rename(
        columns={
            "Loot box name": "BoxName",
            "Max total items": "MaxItems",
            "Min box value": "MinValue",
            "Max box value": "MaxValue",
            "Min scarecity": "MinScarcity",
            "Max scarecity": "MaxScarcity",
        },
        inplace=True,
    )

    # Players sheet: flatten into inventory_df
    players = pd.read_excel(filepath, sheet_name="Players", header=[0, 1])
    records = []
    for player in players.columns.levels[0]:
        if player in ("Players", "Party"):
            continue
        for _, row in players.iterrows():
            item = row.get((player, "Loot"))
            qty = row.get((player, "Qty"))
            if pd.notna(item) and pd.notna(qty):
                records.append({"Player": player, "Item": item, "Qty": int(qty)})
    inventory_df = pd.DataFrame(records)

    # Enrich inventory with item details
    inventory_df = inventory_df.merge(
        items_df[["Item", "Value", "Weight", "Scarcity"]], on="Item", how="left"
    )

    return items_df, boxes_df, players, inventory_df


class LootBoxGeneratorWindow(QMainWindow):
    def __init__(self, data):
        super().__init__()
        self.items_df, self.boxes_df, self.players_tmpl, self.inv_df = data
        self.setWindowTitle("Loot Box Generator")
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        layout = QVBoxLayout(central)
        label = QLabel("Loot Box Generator")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 24px;")
        layout.addWidget(label)
        # Future: add dropdowns, buttons, tables here
        self.setCentralWidget(central)


class PlayerInventoryWindow(QMainWindow):
    def __init__(self, data):
        super().__init__()
        self.items_df, self.boxes_df, self.players_tmpl, self.inv_df = data
        self.setWindowTitle("Player Inventory")
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        layout = QVBoxLayout(central)
        label = QLabel("Player Inventory")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 24px;")
        layout.addWidget(label)
        # Future: add inventory table and controls here
        self.setCentralWidget(central)


if __name__ == "__main__":
    # Load data before creating windows
    data = load_data(EXCEL_FILE)

    app = QApplication(sys.argv)
    # Pass loaded data to both windows
    loot_window = LootBoxGeneratorWindow(data)
    inv_window = PlayerInventoryWindow(data)

    loot_window.show()
    inv_window.show()
    sys.exit(app.exec())
