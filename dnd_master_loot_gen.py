#!/usr/bin/env python3
"""
Loot Master PySide6 App with Real Data and Dynamic Updates

Two windows:
1. Loot Box Generator
2. Player Inventory

Features:
- Load real data from loot_table.xlsx
- Roll loot and "Take" updates inventory immediately
- Auto-update Excel when checkbox is enabled
- Inventory window refreshes dynamically on "Take"

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
    QCheckBox,
    QSlider,
    QDialog,
    QDialogButtonBox,
    QMessageBox,
)
from PySide6.QtCore import Qt

# --- Excel I/O functions ---------------------------------------------------
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "loot_table.xlsx")


def write_inventory(inv_df, players_template, filepath):
    """
    Write back the 'Players' sheet, preserving structure. Merge duplicate items for each player.
    """
    players = [
        p for p in players_template.columns.levels[0] if p not in ("Players", "Party")
    ]
    # Merge duplicate items for each player by summing Qty
    merged = inv_df.groupby(["Player", "Item"], as_index=False).agg(
        {"Qty": "sum", "Value": "first", "Weight": "first", "Scarcity": "first"}
    )
    max_rows = merged.groupby("Player").size().max() if not merged.empty else 0
    # prepare blank DataFrame with same MultiIndex
    new_df = pd.DataFrame(index=range(max_rows), columns=players_template.columns)
    for p in players:
        grp = merged[merged["Player"] == p].reset_index(drop=True)
        for i, row in grp.iterrows():
            new_df.at[i, (p, "Loot")] = row["Item"]
            new_df.at[i, (p, "Qty")] = row["Qty"]
    # write
    with pd.ExcelWriter(
        filepath, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        new_df.to_excel(
            writer, sheet_name="Players"
        )  # allow index column since MultiIndex headers require index


# --- Data loading and logic ------------------------------------------------
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
    # Read players sheet with two header rows manually
    raw = pd.read_excel(filepath, sheet_name="Players", header=None)
    raw = raw.iloc[:, 1:]
    # First two rows are headers
    lvl0 = raw.iloc[0].fillna(method="ffill")
    lvl1 = raw.iloc[1]
    # Data starts from row 2
    data = raw.iloc[2:].reset_index(drop=True)
    # Build MultiIndex columns
    players = data.copy()
    players.columns = pd.MultiIndex.from_arrays([lvl0, lvl1])
    # Flatten inventory records
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
        out.append([r.Scarcity, r.Item, q, val, wgt])
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
        rows.append([r.Scarcity, r.Item, r.Qty, val, wgt])
    return rows


# Shared table builder
def setup_table(table: QTableWidget, headers, rows, action1, action2):
    table.clear()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    table.setAlternatingRowColors(True)
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            it = QTableWidgetItem(str(val))
            it.setFlags(Qt.ItemIsEnabled)
            table.setItem(r_idx, c_idx, it)
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


# --- GUI windows -----------------------------------------------------------
class ExcelOptionsWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Excel Options")
        self.setMinimumSize(250, 100)
        layout = QVBoxLayout()
        self.auto_chk = QCheckBox("Auto-update Excel")
        self.auto_chk.setChecked(True)  # Automatically checked by default
        layout.addWidget(self.auto_chk)
        btns = QDialogButtonBox(QDialogButtonBox.Ok)
        btns.accepted.connect(self.accept)
        layout.addWidget(btns)
        self.setLayout(layout)


class LootBoxGeneratorWindow(QMainWindow):
    def __init__(self, data, excel_options):
        super().__init__()
        self.items, self.boxes, self.players_tmpl, self.inv_df = data
        self.current_rows = []
        self.setWindowTitle("Loot Box Generator")
        self.excel_options = excel_options
        self._build_ui()

    def _build_ui(self):
        w = QWidget()
        v = QVBoxLayout(w)
        # Box selector
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Loot Box:"))
        self.box_combo = QComboBox()
        self.box_combo.addItems(self.boxes.BoxName.tolist())
        h1.addWidget(self.box_combo)
        h1.addStretch()
        roll_btn = QPushButton("Roll")
        roll_btn.setStyleSheet("background:#F55;color:white;")
        roll_btn.clicked.connect(self.on_roll)
        h1.addWidget(roll_btn)
        v.addLayout(h1)
        v.addWidget(self._hr())
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
        take_all_btn.setStyleSheet("background:#F55;color:white;")
        take_all_btn.clicked.connect(self.take_all)
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
        # Excel Options button
        excel_btn = QPushButton("Excel Options")
        excel_btn.clicked.connect(self.show_excel_options)
        v.addWidget(excel_btn)
        self.setCentralWidget(w)
        self.setMinimumSize(500, 327)
        self.resize(500, 327)

    def show_excel_options(self):
        self.excel_options.exec()

    def _hr(self):
        ln = QFrame()
        ln.setFrameShape(QFrame.HLine)
        ln.setFrameShadow(QFrame.Sunken)
        return ln

    def on_roll(self):
        self.current_rows = roll_loot(
            self.box_combo.currentText(), self.items, self.boxes
        )
        self._refresh_table()

    def _refresh_table(self):
        setup_table(
            self.table,
            ["Rarity", "Item", "Qty", "Value", "Weight", "Take", "Drop"],
            self.current_rows,
            self.on_take,
            self.on_drop,
        )
        tw = sum(r[4] for r in self.current_rows)
        tv = sum(r[3] for r in self.current_rows)
        self.wlbl.setText(f"Total Weight: {tw:.1f}")
        self.vlbl.setText(f"Total Value: {tv:.1f}")

    def on_take(self, item):
        for i, r in enumerate(self.current_rows):
            if r[1] == item:
                scar, name, qty, val, wgt = r
                unit_val = round(val / qty, 1)
                unit_wgt = round(wgt / qty, 1)
                self.inv_df.loc[len(self.inv_df)] = {
                    "Player": self.player_combo.currentText(),
                    "Item": name,
                    "Qty": qty,
                    "Value": unit_val,
                    "Weight": unit_wgt,
                    "Scarcity": scar,
                }
                del self.current_rows[i]
                self._refresh_table()
                # Change player in inventory window to match loot window (unless Party is selected)
                if self.inv_window.owner_combo.currentText() != "Party":
                    self.inv_window.owner_combo.setCurrentText(
                        self.player_combo.currentText()
                    )
                    self.inv_window.refresh(self.player_combo.currentText())
                else:
                    self.inv_window.refresh("Party")
                    self.inv_window.owner_combo.setCurrentText("Party")

                # excel update
                if self.excel_options.auto_chk.isChecked():
                    write_inventory(self.inv_df, self.players_tmpl, EXCEL_FILE)
                return

    def on_drop(self, item):
        self.current_rows = [r for r in self.current_rows if r[1] != item]
        self._refresh_table()

    def take_all(self):
        player = self.player_combo.currentText()
        for scar, name, qty, val, wgt in self.current_rows:
            uv = round(val / qty, 1)
            uw = round(wgt / qty, 1)
            self.inv_df.loc[len(self.inv_df)] = {
                "Player": player,
                "Item": name,
                "Qty": qty,
                "Value": uv,
                "Weight": uw,
                "Scarcity": scar,
            }
        self.current_rows = []
        self._refresh_table()
        # Change player in inventory window to match loot window (unless Party is selected)
        if self.inv_window.owner_combo.currentText() != "Party":
            self.inv_window.owner_combo.setCurrentText(self.player_combo.currentText())
        self.inv_window.refresh(self.player_combo.currentText())
        if self.excel_options.auto_chk.isChecked():
            write_inventory(self.inv_df, self.players_tmpl, EXCEL_FILE)

    def closeEvent(self, event):
        QApplication.instance().quit()

    def resizeEvent(self, event):
        pass  # Remove printout


class PlayerInventoryWindow(QMainWindow):
    def __init__(self, data, excel_options):
        super().__init__()
        self.items, self.boxes, self.players_tmpl, self.inv_df = data
        self.setWindowTitle("Player Inventory")
        self.excel_options = excel_options
        self._build_ui()

    def _build_ui(self):
        w = QWidget()
        v = QVBoxLayout(w)
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
        # Add button to manually add items
        add_btn = QPushButton("+")
        add_btn.setFixedWidth(30)
        add_btn.setToolTip("Add item to player inventory")
        add_btn.clicked.connect(self.show_add_item_dialog)
        h1.addWidget(add_btn)
        v.addLayout(h1)
        h2 = QHBoxLayout()
        self.wlbl = QLabel("Total Weight: 0.0")
        h2.addWidget(self.wlbl)
        h2.addStretch()
        self.vlbl = QLabel("Total Value: 0.0")
        h2.addWidget(self.vlbl)
        v.addLayout(h2)
        self.table = QTableWidget()
        v.addWidget(self.table)
        self.setCentralWidget(w)
        self.setMinimumSize(500, 260)
        self.resize(500, 260)
        self.refresh(self.owner_combo.currentText())

    def refresh(self, owner):
        self._current_owner = owner  # Track current owner for drop logic
        rows = get_aggregated(self.inv_df, owner)
        setup_table(
            self.table,
            ["Rarity", "Item", "Qty", "Value", "Weight", "Trade", "Drop"],
            rows,
            self.on_trade,
            self.on_user_drop_or_trade,
        )
        tw = sum(r[4] for r in rows)
        tv = sum(r[3] for r in rows)
        self.wlbl.setText(f"Total Weight: {tw:.1f}")
        self.vlbl.setText(f"Total Value: {tv:.1f}")

    def on_trade(self, item):
        return self.on_user_drop_or_trade(item, drop_or_trade="trade")

    def on_user_drop_or_trade(self, item, drop_or_trade="drop"):
        owner = self.owner_combo.currentText()
        if owner == "Party":
            QMessageBox.information(
                self, "Drop Disabled", "Cannot drop items when 'Party' is selected."
            )
            return
        # Find the row for this item
        df = self.inv_df[self.inv_df["Player"] == owner]
        item_rows = df[df["Item"] == item]
        if item_rows.empty:
            return
        qty = int(item_rows["Qty"].sum())
        if qty > 1:
            # Show dialog with slider
            dlg = QDialog(self)
            if drop_or_trade == "trade":
                dlg.setWindowTitle(f"Trade {item}")
            else:
                dlg.setWindowTitle(f"Drop {item}")
            layout = QVBoxLayout()
            if drop_or_trade == "trade":
                label = QLabel(f"How many '{item}' to trade? (1-{qty})")
            else:
                label = QLabel(f"How many '{item}' to drop? (1-{qty})")
            layout.addWidget(label)
            slider = QSlider(Qt.Horizontal)
            slider.setMinimum(1)
            slider.setMaximum(qty)
            slider.setValue(1)
            layout.addWidget(slider)
            val_label = QLabel("1")
            layout.addWidget(val_label)
            slider.valueChanged.connect(lambda v: val_label.setText(str(v)))
            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            layout.addWidget(buttons)
            dlg.setLayout(layout)
            result = []

            def accept():
                result.append(slider.value())
                dlg.accept()

            buttons.accepted.connect(accept)
            buttons.rejected.connect(dlg.reject)
            if dlg.exec() == QDialog.Accepted and result:
                drop_qty = result[0]
            else:
                return
        else:
            drop_qty = 1
        # Remove drop_qty from inventory
        left = qty - drop_qty
        # Remove all rows for this item/player
        idxs = self.inv_df[
            (self.inv_df["Player"] == owner) & (self.inv_df["Item"] == item)
        ].index
        self.inv_df.drop(idxs, inplace=True)
        if left > 0:
            # Add back the remaining qty as a single row
            row = item_rows.iloc[0].copy()
            row["Qty"] = left
            self.inv_df.loc[len(self.inv_df)] = row
        # Auto-update Excel if enabled (use shared excel_options)
        if self.excel_options.auto_chk.isChecked():
            write_inventory(self.inv_df, self.players_tmpl, EXCEL_FILE)
        self.refresh(owner)

    def show_add_item_dialog(self):
        owner = self.owner_combo.currentText()
        if owner == "Party":
            QMessageBox.information(
                self, "Add Disabled", "Cannot add items when 'Party' is selected."
            )
            return
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Add Item to {owner}")
        layout = QVBoxLayout()
        # Dropdown for items
        item_combo = QComboBox()
        loot_items = self.items["Item"].tolist()
        item_combo.addItems(loot_items)
        layout.addWidget(QLabel("Select Item:"))
        layout.addWidget(item_combo)
        # Slider for quantity
        qty_slider = QSlider(Qt.Horizontal)
        qty_slider.setMinimum(1)

        # Set max based on selected item
        def update_slider_max():
            item = item_combo.currentText()
            max_qty = int(self.items[self.items["Item"] == item]["MaxQty"].iloc[0])
            qty_slider.setMaximum(max_qty)
            qty_slider.setValue(1)

        item_combo.currentTextChanged.connect(update_slider_max)
        update_slider_max()
        layout.addWidget(QLabel("Quantity:"))
        layout.addWidget(qty_slider)
        qty_label = QLabel("1")
        layout.addWidget(qty_label)
        qty_slider.valueChanged.connect(lambda v: qty_label.setText(str(v)))
        # Weight and value display
        value_label = QLabel()
        weight_label = QLabel()

        def update_value_weight():
            item = item_combo.currentText()
            qty = qty_slider.value()
            row = self.items[self.items["Item"] == item].iloc[0]
            total_value = round(row["Value"] * qty, 1)
            total_weight = round(row["Weight"] * qty, 1)
            value_label.setText(f"Total Value: {total_value}")
            weight_label.setText(f"Total Weight: {total_weight}")

        qty_slider.valueChanged.connect(update_value_weight)
        item_combo.currentTextChanged.connect(update_value_weight)
        update_value_weight()
        layout.addWidget(value_label)
        layout.addWidget(weight_label)
        # OK/Cancel
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(buttons)
        dlg.setLayout(layout)
        result = []

        def accept():
            result.append((item_combo.currentText(), qty_slider.value()))
            dlg.accept()

        buttons.accepted.connect(accept)
        buttons.rejected.connect(dlg.reject)
        if dlg.exec() == QDialog.Accepted and result:
            item, qty = result[0]
            row = self.items[self.items["Item"] == item].iloc[0]
            self.inv_df.loc[len(self.inv_df)] = {
                "Player": owner,
                "Item": item,
                "Qty": qty,
                "Value": row["Value"],
                "Weight": row["Weight"],
                "Scarcity": row["Scarcity"],
            }
            if self.excel_options.auto_chk.isChecked():
                write_inventory(self.inv_df, self.players_tmpl, EXCEL_FILE)
            self.refresh(owner)

    def closeEvent(self, event):
        QApplication.instance().quit()

    def resizeEvent(self, event):
        pass  # Remove printout


if __name__ == "__main__":
    data = load_data(EXCEL_FILE)
    app = QApplication(sys.argv)
    excel_options = ExcelOptionsWindow()
    lw = LootBoxGeneratorWindow(data, excel_options)
    iw = PlayerInventoryWindow(data, excel_options)
    lw.inv_window = iw
    lw.show()
    iw.show()
    sys.exit(app.exec())
