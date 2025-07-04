#!/usr/bin/env python3
"""
Loot Master PySide6 App with Real Data

Two windows:
1. Loot Box Generator
2. Player Inventory

Loads real data from ErwinLootTable.xlsx:
- items_df: master loot items
- boxes_df: loot box definitions
- players_template, inv_df: player inventories

Each window populates controls and tables with real data. Buttons still print actions.

Run:
    python loot_master_app.py

Requires:
    PySide6, pandas, openpyxl
"""

import sys
import os
import random
import pandas as pd
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

# --- Data loading and logic ------------------------------------------------
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "ErwinLootTable.xlsx")


def load_data(filepath):
    # Load items
    items = pd.read_excel(filepath, sheet_name="Loot").dropna(subset=["Item"])
    items.rename(
        columns={"Value(GP)": "Value", "Max": "MaxQty", "Item scarecity": "Scarcity"},
        inplace=True,
    )
    # Load boxes
    boxes = pd.read_excel(filepath, sheet_name="Loot box sizes")
    boxes.rename(
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
    # Load players -> flatten inventory
    players = pd.read_excel(filepath, sheet_name="Players", header=[0, 1])
    recs = []
    for p in players.columns.levels[0]:
        if p in ("Players", "Party"):
            continue
        for _, r in players.iterrows():
            it = r.get((p, "Loot"))
            qt = r.get((p, "Qty"))
            if pd.notna(it) and pd.notna(qt):
                recs.append({"Player": p, "Item": it, "Qty": int(qt)})
    inv = pd.DataFrame(recs).merge(
        items[["Item", "Value", "Weight", "Scarcity"]], on="Item", how="left"
    )
    return items, boxes, players, inv


# Roll loot by sampling candidates
def roll_loot(box_name, items_df, boxes_df):
    box = boxes_df[boxes_df.BoxName == box_name].iloc[0]
    c = items_df[
        (items_df.Scarcity >= box.MinScarcity)
        & (items_df.Scarcity <= box.MaxScarcity)
        & (items_df.Value >= box.MinValue)
        & (items_df.Value <= box.MaxValue)
    ]
    n = min(int(box.MaxItems), len(c))
    chosen = c.sample(n)
    out = []
    for _, r in chosen.iterrows():
        q = random.randint(1, int(r.MaxQty))
        val = round(r.Value * q, 1)
        wgt = round(r.Weight * q, 1)
        out.append((r.Scarcity, r.Item, q, val, wgt))
    return out


# Aggregate inventory for a player or party
def get_aggregated(inv_df, player):
    df = inv_df if player == "Party" else inv_df[inv_df.Player == player]
    if df.empty:
        return []
    agg = df.groupby(["Item", "Scarcity"]).agg({"Qty": "sum"}).reset_index()
    rows = []
    for _, r in agg.iterrows():
        base = df[df.Item == r.Item].iloc[0]
        val = round(r.Qty * base.Value, 1)
        wgt = round(r.Qty * base.Weight, 1)
        rows.append((r.Scarcity, r.Item, r.Qty, val, wgt))
    return rows


# --- Shared table builder --------------------------------------------------
def setup_table(table: QTableWidget, headers, rows, action1, action2):
    table.clear()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    table.setAlternatingRowColors(True)
    table.setStyleSheet(
        "QHeaderView::section{background:#aaa;}QTableWidget{gridline-color:#666}"
    )
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            item = QTableWidgetItem(str(val))
            item.setFlags(Qt.ItemIsEnabled)
            table.setItem(r_idx, c_idx, item)
        btn1 = QPushButton(headers[-2])
        btn1.setStyleSheet("background:#888;color:white;border-radius:4px;")
        btn1.clicked.connect(lambda _, it=row[1]: action1(it))
        table.setCellWidget(r_idx, len(headers) - 2, btn1)
        btn2 = QPushButton(headers[-1])
        btn2.setStyleSheet("background:#F55;color:white;border-radius:4px;")
        btn2.clicked.connect(lambda _, it=row[1]: action2(it))
        table.setCellWidget(r_idx, len(headers) - 1, btn2)
    table.resizeColumnsToContents()
    table.horizontalHeader().setStretchLastSection(True)


# --- Loot Box Generator Window ---------------------------------------------
class LootBoxGeneratorWindow(QMainWindow):
    def __init__(self, data):
        super().__init__()
        self.items, self.boxes, self.players_tmpl, self.inv_df = data
        self.setWindowTitle("Loot Box Generator")
        self._build_ui()

    def _build_ui(self):
        w = QWidget()
        v = QVBoxLayout(w)
        # Controls
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Loot Box:"))
        self.box_combo = QComboBox()
        self.box_combo.addItems(self.boxes.BoxName.tolist())
        h1.addWidget(self.box_combo)
        h1.addStretch()
        roll_btn = QPushButton("Roll")
        roll_btn.setStyleSheet(
            "background:#F55;color:white;padding:6px;border-radius:4px;"
        )
        roll_btn.clicked.connect(self.on_roll)
        h1.addWidget(roll_btn)
        v.addLayout(h1)
        # Separator line
        ln = QFrame()
        ln.setFrameShape(QFrame.HLine)
        ln.setFrameShadow(QFrame.Sunken)
        v.addWidget(ln)
        # Player selector
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("Player:"))
        players = [
            p
            for p in self.players_tmpl.columns.levels[0]
            if p not in ("Players", "Party")
        ]
        self.player_combo = QComboBox()
        self.player_combo.addItems(players)
        h2.addWidget(self.player_combo)
        h2.addStretch()
        take_all_btn = QPushButton("Take All")
        take_all_btn.setStyleSheet(
            "background:#F55;color:white;padding:6px;border-radius:4px;"
        )
        take_all_btn.clicked.connect(lambda: print("Take All"))
        h2.addWidget(take_all_btn)
        v.addLayout(h2)
        # Totals
        h3 = QHBoxLayout()
        self.wlbl = QLabel("Total Weight: 0.0")
        h3.addWidget(self.wlbl)
        h3.addStretch()
        self.vlbl = QLabel("Total Value: 0.0")
        h3.addWidget(self.vlbl)
        v.addLayout(h3)
        # Table
        self.table = QTableWidget()
        v.addWidget(self.table)
        self.setCentralWidget(w)

    def on_roll(self):
        box = self.box_combo.currentText()
        print(f"Roll {box}")
        rows = roll_loot(box, self.items, self.boxes)
        setup_table(
            self.table,
            ["Rarity", "Item", "Qty", "Value", "Weight", "Take", "Drop"],
            rows,
            self.on_take,
            self.on_drop,
        )
        tw = sum(r[4] for r in rows)
        tv = sum(r[3] for r in rows)
        self.wlbl.setText(f"Total Weight: {tw:.1f}")
        self.vlbl.setText(f"Total Value: {tv:.1f}")

    def on_take(self, item):
        print(f"Take {item}")

    def on_drop(self, item):
        print(f"Drop {item}")


# --- Player Inventory Window ----------------------------------------------
class PlayerInventoryWindow(QMainWindow):
    def __init__(self, data):
        super().__init__()
        self.items, self.boxes, self.players_tmpl, self.inv_df = data
        self.setWindowTitle("Player Inventory")
        self._build_ui()

    def _build_ui(self):
        w = QWidget()
        v = QVBoxLayout(w)
        # Owner selector
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Player:"))
        owners = [
            p
            for p in self.players_tmpl.columns.levels[0]
            if p not in ("Players", "Party")
        ] + ["Party"]
        self.owner_combo = QComboBox()
        self.owner_combo.addItems(owners)
        self.owner_combo.currentTextChanged.connect(self.refresh)
        h1.addWidget(self.owner_combo)
        h1.addStretch()
        v.addLayout(h1)
        # Totals
        h2 = QHBoxLayout()
        self.wlbl = QLabel("Total Weight: 0.0")
        h2.addWidget(self.wlbl)
        h2.addStretch()
        self.vlbl = QLabel("Total Value: 0.0")
        h2.addWidget(self.vlbl)
        v.addLayout(h2)
        # Table
        self.table = QTableWidget()
        v.addWidget(self.table)
        self.setCentralWidget(w)
        # initial load
        self.refresh(self.owner_combo.currentText())

    def refresh(self, owner):
        print(f"Refresh {owner}")
        rows = get_aggregated(self.inv_df, owner)
        setup_table(
            self.table,
            ["Rarity", "Item", "Qty", "Value", "Weight", "Trade", "Drop"],
            rows,
            self.on_trade,
            self.on_drop,
        )
        tw = sum(r[4] for r in rows)
        tv = sum(r[3] for r in rows)
        self.wlbl.setText(f"Total Weight: {tw:.1f}")
        self.vlbl.setText(f"Total Value: {tv:.1f}")

    def on_trade(self, item):
        print(f"Trade {item}")

    def on_drop(self, item):
        print(f"Drop {item}")


# --- Main entry -----------------------------------------------------------
if __name__ == "__main__":
    data = load_data(EXCEL_FILE)
    app = QApplication(sys.argv)
    lw = LootBoxGeneratorWindow(data)
    iw = PlayerInventoryWindow(data)
    lw.show()
    iw.show()
    sys.exit(app.exec())
