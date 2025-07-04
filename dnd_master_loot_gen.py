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
    Handles OS/I/O errors with user-friendly dialogs.
    """
    import errno

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
    try:
        with pd.ExcelWriter(
            filepath, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            new_df.to_excel(
                writer, sheet_name="Players"
            )  # allow index column since MultiIndex headers require index
    except PermissionError as e:
        QMessageBox.critical(
            None,
            "Excel File Locked or Permission Denied",
            f"Cannot write to Excel file.\n\nReason: {str(e)}\n\nPlease close the file in Excel or check your permissions.",
        )
    except OSError as e:
        if e.errno == errno.ENOSPC:
            QMessageBox.critical(None, "Disk Full", "Saving failed: Disk is full.")
        elif e.errno == errno.EINVAL:
            QMessageBox.critical(
                None, "Invalid File Path", f"Invalid file path: {filepath}"
            )
        else:
            QMessageBox.critical(
                None, "File Write Error", f"Error writing to Excel file:\n{str(e)}"
            )
    except Exception as e:
        QMessageBox.critical(
            None, "Unknown Error", f"Unexpected error writing Excel file:\n{str(e)}"
        )


# --- Data loading and logic ------------------------------------------------
def load_data(filepath):
    """
    Load all data from Excel, with robust error handling for file format issues.
    Returns: items, boxes, players, inv
    """
    import shutil

    # --- Helper: show error and offer to regenerate or rename ---
    def show_format_error(msg, xl_ref=None):
        ret = QMessageBox.question(
            None,
            "Excel Format Error",
            msg
            + "\n\nWould you like to regenerate the file?\n\nYes: Overwrite with a new blank file.\nNo: Rename current file to 'loot_table_error.xlsx' and create a new blank one.",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
        )
        if ret == QMessageBox.Yes:
            create_default_loot_excel(filepath)
            return True
        elif ret == QMessageBox.No:
            # Ensure all file handles are closed before renaming
            if xl_ref is not None:
                try:
                    xl_ref.close()
                except Exception:
                    pass
            import gc

            gc.collect()  # Extra safety to release file handles
            error_path = os.path.join(
                os.path.dirname(filepath), "loot_table_error.xlsx"
            )
            if os.path.exists(error_path):
                ow = QMessageBox.question(
                    None,
                    "Error File Exists",
                    f"'{error_path}' already exists. Overwrite it?",
                    QMessageBox.Yes | QMessageBox.Cancel,
                )
                if ow != QMessageBox.Yes:
                    sys.exit(1)
            try:
                shutil.move(filepath, error_path)
            except Exception as e:
                QMessageBox.critical(
                    None, "File Rename Error", f"Could not rename file: {str(e)}"
                )
                sys.exit(1)
            create_default_loot_excel(filepath)
            return True
        else:
            sys.exit(1)

    # --- Helper: check required columns ---
    def check_columns(df, required, sheet, xl_ref=None):
        missing = [c for c in required if c not in df.columns]
        if missing:
            if show_format_error(
                f"Missing columns in '{sheet}' sheet: {missing}", xl_ref=xl_ref
            ):
                raise Exception("Regenerated file")

    # --- Helper: validate all rows in a DataFrame for numeric columns ---
    def validate_numeric_columns(df, columns, sheet_name):
        for col in columns:
            if col not in df.columns:
                QMessageBox.critical(
                    None,
                    f"Missing Column in {sheet_name} Sheet",
                    f"Column '{col}' is missing from the {sheet_name} sheet.\nPlease fix the sheet in Excel and reload.",
                )
                sys.exit(1)
            # Check all rows for non-numeric values
            for idx, val in df[col].items():
                if not pd.api.types.is_number(val):
                    item_name = (
                        df.loc[idx, "Item"] if "Item" in df.columns else str(idx)
                    )
                    QMessageBox.critical(
                        None,
                        "Invalid Data in Excel",
                        f"Non-numeric value in column '{col}' for item '{item_name}' in the {sheet_name} sheet.\nPlease fix this value in Excel and reload.",
                    )
                    sys.exit(1)
            # Coerce to numeric if all are valid
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # --- Try to load all sheets, handle missing/corrupt ---
    try:
        xl = pd.ExcelFile(filepath)
    except Exception as e:
        if show_format_error(f"Could not open Excel file.\nReason: {str(e)}"):
            raise Exception("Regenerated file")

    required_sheets = ["Loot", "Loot box sizes", "Players"]
    missing_sheets = [s for s in required_sheets if s not in xl.sheet_names]
    if missing_sheets:
        if show_format_error(
            f"Excel file is missing required sheets: {missing_sheets}", xl_ref=xl
        ):
            raise Exception("Regenerated file")

    # --- Loot sheet ---
    try:
        items = xl.parse("Loot").dropna(subset=["Item"])
    except Exception as e:
        if show_format_error(
            f"Could not read 'Loot' sheet.\nReason: {str(e)}", xl_ref=xl
        ):
            raise Exception("Regenerated file")
    items.rename(
        columns={"Value(GP)": "Value", "Max": "MaxQty", "Item scarecity": "Scarcity"},
        inplace=True,
    )
    check_columns(
        items, ["Item", "Value", "MaxQty", "Weight", "Scarcity"], "Loot", xl_ref=xl
    )
    # Validate all rows in Loot sheet for numeric columns
    validate_numeric_columns(items, ["Value", "MaxQty", "Weight", "Scarcity"], "Loot")

    # --- Loot box sizes sheet ---
    try:
        boxes = xl.parse("Loot box sizes")
    except Exception as e:
        if show_format_error(
            f"Could not read 'Loot box sizes' sheet.\nReason: {str(e)}", xl_ref=xl
        ):
            raise Exception("Regenerated file")
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
    check_columns(
        boxes,
        ["BoxName", "MaxItems", "MinValue", "MaxValue", "MinScarcity", "MaxScarcity"],
        "Loot box sizes",
        xl_ref=xl,
    )
    # Validate all rows in Loot box sizes sheet for numeric columns
    validate_numeric_columns(
        boxes,
        ["MaxItems", "MinValue", "MaxValue", "MinScarcity", "MaxScarcity"],
        "Loot box sizes",
    )

    # --- Players sheet: check MultiIndex header ---
    try:
        raw = xl.parse("Players", header=None)
    except Exception as e:
        if show_format_error(
            f"Could not read 'Players' sheet.\nReason: {str(e)}", xl_ref=xl
        ):
            raise Exception("Regenerated file")
    if raw.shape[0] < 3 or raw.shape[1] < 3:
        if show_format_error(
            "'Players' sheet format is invalid (not enough rows/columns).", xl_ref=xl
        ):
            raise Exception("Regenerated file")
    raw = raw.iloc[:, 1:]
    lvl0 = raw.iloc[0].fillna(method="ffill")
    lvl1 = raw.iloc[1]
    if lvl0.isnull().any() or lvl1.isnull().any():
        if show_format_error(
            "'Players' sheet header rows contain missing values.", xl_ref=xl
        ):
            raise Exception("Regenerated file")
    data = raw.iloc[2:].reset_index(drop=True)
    try:
        players = data.copy()
        players.columns = pd.MultiIndex.from_arrays([lvl0, lvl1])
    except Exception as e:
        if show_format_error(
            f"'Players' sheet header format error: {str(e)}", xl_ref=xl
        ):
            raise Exception("Regenerated file")

    # --- Validate all player Qty columns in Players sheet ---
    for player in players.columns.levels[0]:
        if player in ("Players", "Party"):
            continue
        col = (player, "Qty")
        if col not in players.columns:
            QMessageBox.critical(
                None,
                "Missing Column in Players Sheet",
                f"Column 'Qty' for player '{player}' is missing from the Players sheet.\nPlease fix the sheet in Excel and reload.",
            )
            sys.exit(1)
        for idx, val in players[col].items():
            if not pd.api.types.is_number(val):
                loot_val = (
                    players[(player, "Loot")][idx]
                    if (player, "Loot") in players.columns
                    else str(idx)
                )
                QMessageBox.critical(
                    None,
                    "Invalid Data in Excel",
                    f"Non-numeric value in column 'Qty' for player '{player}', item '{loot_val}' (row {idx+3}) in the Players sheet.\nPlease fix this value in Excel and reload.",
                )
                sys.exit(1)
        players[col] = pd.to_numeric(players[col], errors="coerce")

    # --- Validate inventory items exist in loot table ---
    recs = []
    for p in players.columns.levels[0]:
        if p in ("Players", "Party"):
            continue
        for _, r in players.iterrows():
            it = r.get((p, "Loot"))
            qt = r.get((p, "Qty"))
            if pd.notna(it) and pd.notna(qt):
                recs.append({"Player": p, "Item": it, "Qty": int(qt)})
    inv = pd.DataFrame(recs)
    missing_items = set(inv["Item"]) - set(items["Item"])
    if missing_items:
        msg = (
            f"The following inventory items are missing from the loot table:\n"
            f"{list(missing_items)}\n\nPlease add them to the 'Loot' sheet and click Retry."
        )
        while True:
            ret = QMessageBox.question(
                None,
                "Missing Inventory Items",
                msg,
                QMessageBox.Retry | QMessageBox.Abort,
            )
            if ret == QMessageBox.Retry:
                xl.close()
                xl = pd.ExcelFile(filepath)
                items = xl.parse("Loot").dropna(subset=["Item"])
                items.rename(
                    columns={
                        "Value(GP)": "Value",
                        "Max": "MaxQty",
                        "Item scarecity": "Scarcity",
                    },
                    inplace=True,
                )
                if set(inv["Item"]) - set(items["Item"]):
                    continue
                break
            else:
                xl.close()
                sys.exit(1)
    xl.close()
    # --- Validate numeric columns in inventory (Players) ---
    numeric_cols = ["Qty"]
    for col in numeric_cols:
        if col not in inv.columns:
            QMessageBox.critical(
                None,
                "Missing Column in Players Sheet",
                f"Column '{col}' is missing from the Players sheet.\nPlease fix the sheet in Excel and reload.",
            )
            sys.exit(1)
        for idx, val in inv[col].items():
            if not pd.api.types.is_number(val):
                item_name = inv.loc[idx, "Item"] if "Item" in inv.columns else str(idx)
                QMessageBox.critical(
                    None,
                    "Invalid Data in Excel",
                    f"Non-numeric value in column '{col}' for item '{item_name}' in the Players sheet.\nPlease fix this value in Excel and reload.",
                )
                sys.exit(1)
        inv[col] = pd.to_numeric(inv[col], errors="coerce")
    # Merge loot info
    inv = inv.merge(
        items[["Item", "Value", "Weight", "Scarcity"]], on="Item", how="left"
    )
    # Validate loot columns are numeric and present in merged inventory
    validate_numeric_columns(inv, ["Value", "Weight", "Scarcity"], "Loot (Inventory)")
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


# Create default loot Excel file with predefined structure
def create_default_loot_excel(filepath: str):
    """
    Create a new Excel file at `filepath` with three sheets formatted for Loot Master App.

    1) Loot
       Columns: Item, Description, Value(GP), Max, Weight, Item scarecity
       - Gold Coin
       - Silver Coin
       - Jewelry Box

    2) Loot box sizes
       Columns: Loot box name, Max total items, Min box value, Max box value, Min scarecity, Max scarecity
       - Small Chest
       - Medium Chest
       - Large Chest

    3) Players
       MultiIndex columns: level 0 = [Player 1, Player 2], level 1 = [Loot, Qty]
       One row: each player starts with 10 Gold Coins.
    """
    # 1) Loot sheet
    loot_df = pd.DataFrame(
        [
            {
                "Item": "Gold Coin",
                "Description": "Standard gold currency",
                "Value(GP)": 1.0,
                "Max": 100,
                "Weight": 0.02,
                "Item scarecity": 1,
            },
            {
                "Item": "Silver Coin",
                "Description": "Standard silver currency",
                "Value(GP)": 0.1,
                "Max": 200,
                "Weight": 0.01,
                "Item scarecity": 2,
            },
            {
                "Item": "Jewelry Box",
                "Description": "Small ornate jewelry box",
                "Value(GP)": 50.0,
                "Max": 1,
                "Weight": 1.0,
                "Item scarecity": 5,
            },
        ]
    )

    # 2) Loot box sizes sheet
    boxes_df = pd.DataFrame(
        [
            {
                "Loot box name": "Small Chest",
                "Max total items": 3,
                "Min box value": 1.0,
                "Max box value": 20.0,
                "Min scarecity": 1,
                "Max scarecity": 3,
            },
            {
                "Loot box name": "Medium Chest",
                "Max total items": 5,
                "Min box value": 10.0,
                "Max box value": 50.0,
                "Min scarecity": 1,
                "Max scarecity": 5,
            },
            {
                "Loot box name": "Large Chest",
                "Max total items": 10,
                "Min box value": 20.0,
                "Max box value": 200.0,
                "Min scarecity": 1,
                "Max scarecity": 6,
            },
        ]
    )

    # 3) Players sheet with MultiIndex columns
    players_cols = pd.MultiIndex.from_product(
        [["Player 1", "Player 2"], ["Loot", "Qty"]], names=[None, None]
    )
    # single row: both players start with 10 Gold Coins
    players_df = pd.DataFrame(
        [["Gold Coin", 10, "Gold Coin", 10]], columns=players_cols
    )

    # Write to Excel
    try:
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            loot_df.to_excel(writer, sheet_name="Loot", index=False)
            players_df.to_excel(writer, sheet_name="Players")
            boxes_df.to_excel(writer, sheet_name="Loot box sizes", index=False)
    except PermissionError as e:
        QMessageBox.critical(
            None,
            "Excel File Locked or Permission Denied",
            f"Cannot create Excel file.\n\nReason: {str(e)}\n\nPlease close the file in Excel or check your permissions.",
        )
    except OSError as e:
        import errno

        if e.errno == errno.ENOSPC:
            QMessageBox.critical(
                None, "Disk Full", "Creating file failed: Disk is full."
            )
        elif e.errno == errno.EINVAL:
            QMessageBox.critical(
                None, "Invalid File Path", f"Invalid file path: {filepath}"
            )
        else:
            QMessageBox.critical(
                None, "File Create Error", f"Error creating Excel file:\n{str(e)}"
            )
    except Exception as e:
        QMessageBox.critical(
            None, "Unknown Error", f"Unexpected error creating Excel file:\n{str(e)}"
        )


# --- GUI windows -----------------------------------------------------------
class ExcelOptionsWindow(QDialog):
    def __init__(self, parent=None, reload_callback=None, write_callback=None):
        super().__init__(parent)
        self.setWindowTitle("Excel Options")
        self.setMinimumSize(250, 150)
        layout = QVBoxLayout()
        self.auto_chk = QCheckBox("Auto-update Excel")
        self.auto_chk.setChecked(True)  # Automatically checked by default
        layout.addWidget(self.auto_chk)
        # Read and Write buttons
        btn_layout = QHBoxLayout()
        self.read_btn = QPushButton("Read")
        self.write_btn = QPushButton("Write")
        btn_layout.addWidget(self.read_btn)
        btn_layout.addWidget(self.write_btn)
        layout.addLayout(btn_layout)

        # Disable buttons if auto-update is checked
        def update_btns():
            enabled = not self.auto_chk.isChecked()
            self.read_btn.setEnabled(enabled)
            self.write_btn.setEnabled(enabled)

        self.auto_chk.stateChanged.connect(update_btns)
        update_btns()
        btns = QDialogButtonBox(QDialogButtonBox.Ok)
        btns.accepted.connect(self.accept)
        layout.addWidget(btns)
        self.setLayout(layout)
        self.reload_callback = reload_callback
        self.write_callback = write_callback
        self.read_btn.clicked.connect(self.on_read)
        self.write_btn.clicked.connect(self.on_write)

    def on_read(self):
        if self.reload_callback:
            self.reload_callback()

    def on_write(self):
        if self.write_callback:
            self.write_callback()


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
    # Check if Excel file exists, if not, create it (with error handling)
    if not os.path.exists(EXCEL_FILE):
        try:
            create_default_loot_excel(EXCEL_FILE)
        except Exception as e:
            QMessageBox.critical(
                None, "Excel File Error", f"Failed to create Excel file:\n{str(e)}"
            )
            sys.exit(1)
    app = QApplication(sys.argv)

    def reload_all():
        new_data = load_data(EXCEL_FILE)
        lw.items, lw.boxes, lw.players_tmpl, lw.inv_df = new_data
        iw.items, iw.boxes, iw.players_tmpl, iw.inv_df = new_data
        lw.box_combo.clear()
        lw.box_combo.addItems(lw.boxes.BoxName.tolist())
        players = [
            p
            for p in lw.players_tmpl.columns.levels[0]
            if p not in ("Players", "Party")
        ]
        lw.player_combo.clear()
        lw.player_combo.addItems(players)
        owners = players + ["Party"]
        iw.owner_combo.clear()
        iw.owner_combo.addItems(owners)
        lw._refresh_table()
        iw.refresh(iw.owner_combo.currentText())

    def write_all():
        # Write all player inventories to Excel using the current data
        write_inventory(iw.inv_df, iw.players_tmpl, EXCEL_FILE)

    excel_options = ExcelOptionsWindow(
        reload_callback=reload_all, write_callback=write_all
    )
    data = load_data(EXCEL_FILE)
    lw = LootBoxGeneratorWindow(data, excel_options)
    iw = PlayerInventoryWindow(data, excel_options)
    lw.inv_window = iw
    lw.show()
    iw.show()
    sys.exit(app.exec())
