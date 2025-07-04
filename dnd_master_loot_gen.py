#!/usr/bin/env python3
"""
Loot Master PySide6 App with Styled Tables and Dummy Data

Two windows:
1. Loot Box Generator
2. Player Inventory

Each screen has:
- Header controls (comboboxes, buttons)
- Totals display
- Styled table with columns [Rarity, Item, Qty, Value, Weight, Take, Drop]
- Dummy rows
- Buttons in cells print actions when clicked

Run:
    python loot_master_app.py

Requires:
    PySide6
"""

import sys
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLabel,
    QComboBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QHBoxLayout,
    QFrame,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor


# --- Shared table builder --------------------------------------------------
def setup_table(
    table: QTableWidget,
    headers: list[str],
    rows: list[tuple],
    take_callback,
    drop_callback,
):
    """
    Configure QTableWidget with given headers and row data.
    Last two columns are 'Take' and 'Drop' with buttons.
    """
    table.clear()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    table.setAlternatingRowColors(True)
    table.setStyleSheet(
        "QHeaderView::section { background: #aaa; padding: 4px; }"
        "QTableWidget { gridline-color: #666; }"
    )
    # Fill rows
    for r, row in enumerate(rows):
        # Data columns
        for c, value in enumerate(row):
            item = QTableWidgetItem(str(value))
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            table.setItem(r, c, item)
        # Take button
        btn_take = QPushButton("✓")
        btn_take.setStyleSheet("background: #888; color: white; border-radius: 4px;")
        btn_take.clicked.connect(lambda _, it=row[1]: take_callback(it))
        table.setCellWidget(r, len(headers) - 2, btn_take)
        # Drop button
        btn_drop = QPushButton("X")
        btn_drop.setStyleSheet("background: #F55; color: white; border-radius: 4px;")
        btn_drop.clicked.connect(lambda _, it=row[1]: drop_callback(it))
        table.setCellWidget(r, len(headers) - 1, btn_drop)
    table.resizeColumnsToContents()
    table.horizontalHeader().setStretchLastSection(True)


# --- Loot Box Generator Window ---------------------------------------------
class LootBoxGeneratorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Loot Box Generator")
        self._build_ui()

    def _build_ui(self):
        w = QWidget()
        v = QVBoxLayout(w)

        # Top controls
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Loot Box:"))
        self.box_combo = QComboBox()
        self.box_combo.addItems(["Large Chest", "Small Bag"])
        h1.addWidget(self.box_combo)
        h1.addStretch()
        roll_btn = QPushButton("Roll")
        roll_btn.setStyleSheet(
            "background: #F55; color: white; padding: 6px; border-radius: 4px;"
        )
        roll_btn.clicked.connect(self.on_roll)
        h1.addWidget(roll_btn)
        v.addLayout(h1)

        # Separator line
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        v.addWidget(line)

        # Player controls
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("Player:"))
        self.player_combo = QComboBox()
        self.player_combo.addItems(["MrTinMan"])
        h2.addWidget(self.player_combo)
        h2.addStretch()
        takeall_btn = QPushButton("Take All")
        takeall_btn.setStyleSheet(
            "background: #F55; color: white; padding: 6px; border-radius: 4px;"
        )
        takeall_btn.clicked.connect(lambda: print("Take All pressed"))
        h2.addWidget(takeall_btn)
        v.addLayout(h2)

        # Totals
        h3 = QHBoxLayout()
        self.weight_lbl = QLabel("Total Weight: 7.3")
        h3.addWidget(self.weight_lbl)
        h3.addStretch()
        self.value_lbl = QLabel("Total Value: 2210")
        h3.addWidget(self.value_lbl)
        v.addLayout(h3)

        # Table
        self.table = QTableWidget()
        v.addWidget(self.table)

        # Populate dummy
        headers = ["Rarity", "Item", "Qty", "Value", "Weight", "Take", "Drop"]
        dummy = [
            (2, "Gold Coin", 10, 10, 1),
            (3, "Sunstone", 2, 150, 4),
            (6, "Jeweled Lockbox", 1, 850, 2),
            (4, "True Ice Necklace", 3, 1200, 0.3),
        ]
        setup_table(self.table, headers, dummy, self.on_take, self.on_drop)

        self.setCentralWidget(w)

    def on_roll(self):
        print(f"Roll pressed (box={self.box_combo.currentText()})")

    def on_take(self, item):
        print(f"Take {item}")

    def on_drop(self, item):
        print(f"Drop {item}")


# --- Player Inventory Window ----------------------------------------------
class PlayerInventoryWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Player Inventory")
        self._build_ui()

    def _build_ui(self):
        w = QWidget()
        v = QVBoxLayout(w)

        # Player selector
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Player:"))
        self.owner_combo = QComboBox()
        self.owner_combo.addItems(["MrTinMan", "Party"])
        h1.addWidget(self.owner_combo)
        h1.addStretch()
        v.addLayout(h1)

        # Totals
        h2 = QHBoxLayout()
        self.weight_lbl = QLabel("Total Weight: 7.3")
        h2.addWidget(self.weight_lbl)
        h2.addStretch()
        self.value_lbl = QLabel("Total Value: 2210")
        h2.addWidget(self.value_lbl)
        v.addLayout(h2)

        # Table
        self.table = QTableWidget()
        v.addWidget(self.table)

        # Dummy inventory rows
        headers = ["Rarity", "Item", "Qty", "Value", "Weight", "Trade", "Drop"]
        dummy = [
            (2, "Gold Coin", 10, 10, 1),
            (3, "Sunstone", 2, 150, 4),
            (6, "Jeweled Lockbox", 1, 850, 2),
            (4, "True Ice Necklace", 3, 1200, 0.3),
        ]
        setup_table(self.table, headers, dummy, self.on_trade, self.on_drop)

        self.setCentralWidget(w)

    def on_trade(self, item):
        print(f"Trade {item}")

    def on_drop(self, item):
        print(f"Drop {item}")


# --- Main entry -----------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    loot_win = LootBoxGeneratorWindow()
    inv_win = PlayerInventoryWindow()
    loot_win.show()
    inv_win.show()
    sys.exit(app.exec())
