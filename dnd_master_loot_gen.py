#!/usr/bin/env python3
"""
Minimal PySide6 App with Two Windows

This application opens two separate windows:
1. Loot Box Generator
2. Player Inventory

Run:
    python minimal_pyside_app.py

Requires:
    PySide6
"""

import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel, QWidget, QVBoxLayout
from PySide6.QtCore import Qt


class LootBoxGeneratorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Loot Box Generator")
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        layout = QVBoxLayout(central)
        label = QLabel("Loot Box Generator")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 24px;")
        layout.addWidget(label)
        self.setCentralWidget(central)


class PlayerInventoryWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Player Inventory")
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        layout = QVBoxLayout(central)
        label = QLabel("Player Inventory")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 24px;")
        layout.addWidget(label)
        self.setCentralWidget(central)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Create and show both windows
    loot_window = LootBoxGeneratorWindow()
    inv_window = PlayerInventoryWindow()
    loot_window.show()
    inv_window.show()

    sys.exit(app.exec())
