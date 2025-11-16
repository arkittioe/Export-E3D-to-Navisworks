# -*- coding: utf-8 -*-

import sys
import os
import json
from pathlib import Path

# Import necessary components from PyQt6
from PyQt6.QtGui import QIcon, QPixmap, QFont, QDesktopServices
from PyQt6.QtCore import Qt, QSettings, QUrl, QTime, QPropertyAnimation, QRect
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox,
    QFileDialog, QMessageBox, QHBoxLayout, QGroupBox, QStatusBar,
    QDialog, QTextBrowser, QTimeEdit, QListWidget, QListWidgetItem
)


class SidePanel(QWidget):
    """
    Side panel widget for managing default object lists per project.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(300)
        self.setup_ui()

    def setup_ui(self):
        """Initialize the side panel UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Header
        header = QLabel("📋 Default Object List")
        header.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(header)

        # Info label
        info = QLabel("Objects used when SITE.txt is not specified:")
        info.setWordWrap(True)
        info.setStyleSheet("font-size: 11px; color: gray;")
        layout.addWidget(info)

        # List widget
        self.list_widget = QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        layout.addWidget(self.list_widget)

        # Button layout
        btn_layout = QHBoxLayout()

        self.btn_add = QPushButton("➕ Add")
        self.btn_add.clicked.connect(self.add_item)
        btn_layout.addWidget(self.btn_add)

        self.btn_remove = QPushButton("➖ Remove")
        self.btn_remove.clicked.connect(self.remove_item)
        btn_layout.addWidget(self.btn_remove)

        self.btn_clear = QPushButton("🗑️ Clear")
        self.btn_clear.clicked.connect(self.clear_all)
        btn_layout.addWidget(self.btn_clear)

        layout.addLayout(btn_layout)

    def add_item(self):
        """Add a new item to the list."""
        from PyQt6.QtWidgets import QInputDialog
        text, ok = QInputDialog.getText(self, "Add Object", "Enter object path:")
        if ok and text.strip():
            self.list_widget.addItem(text.strip())

    def remove_item(self):
        """Remove selected item from the list."""
        current_row = self.list_widget.currentRow()
        if current_row >= 0:
            self.list_widget.takeItem(current_row)

    def clear_all(self):
        """Clear all items."""
        reply = QMessageBox.question(
            self, "Clear List",
            "Are you sure you want to clear all items?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.list_widget.clear()

    def set_items(self, items):
        """Set list items from a list."""
        self.list_widget.clear()
        for item in items:
            if item.strip():
                self.list_widget.addItem(item.strip())

    def get_items(self):
        """Get all items as a list."""
        return [self.list_widget.item(i).text()
                for i in range(self.list_widget.count())]


class AppGUI(QMainWindow):
    """
    Main application GUI class for generating AVEVA E3D export files.
    """

    def __init__(self):
        super().__init__()
        self.settings = QSettings('3DDesignMagic', 'E3DExporter')
        self.side_panel = None  # اضافه کنید
        self.initUI()
        self.load_theme()
        self.connect_signals()

    def initUI(self):
        """
        Initializes the user interface with improved structure and grouping.
        """
        self.setWindowTitle("Export E3D To Navis App")
        self.setFixedSize(700, 800)  # ارتفاع رو کمی بیشتر کردم برای جا دادن time picker

        # Central Widget and Layout
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 10)
        main_layout.setSpacing(15)

        # --- Theme Selection ---
        theme_layout = QHBoxLayout()
        theme_label = QLabel("🎨 Interface Theme:")
        self.combo_mode = QComboBox()
        self.combo_mode.addItems(["Light", "Dark"])
        self.combo_mode.setCurrentText("Dark")
        self.combo_mode.setToolTip("Select your preferred theme")
        theme_layout.addWidget(theme_label)
        theme_layout.addWidget(self.combo_mode)
        theme_layout.addStretch()
        main_layout.addLayout(theme_layout)

        # --- AVEVA Configuration Group ---
        aveva_group = QGroupBox("AVEVA E3D Configuration")
        aveva_group.setStyleSheet("QGroupBox { font-weight: bold; padding-top: 10px; }")
        aveva_layout = QGridLayout()
        aveva_layout.setHorizontalSpacing(15)
        aveva_layout.setVerticalSpacing(10)

        self.labels = {}
        self.line_edits = {}
        self.buttons = {}

        # AVEVA Folder
        self.labels["aveva_folder"] = QLabel("📁 AVEVA Folder:")
        self.line_edits["aveva_folder"] = QLineEdit("C:/Program Files (x86)/AVEVA/Everything3D3.1")
        self.line_edits["aveva_folder"].setToolTip("Path to AVEVA E3D installation folder")
        self.buttons["browse_aveva"] = QPushButton("📂 Browse")
        self.buttons["browse_aveva"].setObjectName("browse")
        self.buttons["browse_aveva"].setToolTip("Browse for AVEVA folder")

        aveva_layout.addWidget(self.labels["aveva_folder"], 0, 0)
        aveva_layout.addWidget(self.line_edits["aveva_folder"], 0, 1)
        aveva_layout.addWidget(self.buttons["browse_aveva"], 0, 2)

        aveva_group.setLayout(aveva_layout)
        main_layout.addWidget(aveva_group)

        # --- Project Information Group ---
        project_group = QGroupBox("Project Information")
        project_group.setStyleSheet("QGroupBox { font-weight: bold; padding-top: 10px; }")
        project_layout = QGridLayout()
        project_layout.setHorizontalSpacing(15)
        project_layout.setVerticalSpacing(10)

        # Project Code (ComboBox با امکان ویرایش)
        self.labels["proj_code"] = QLabel("🏗️ Project Code:")
        self.combo_proj_code = QComboBox()
        self.combo_proj_code.setEditable(True)  # امکان تایپ دستی
        self.combo_proj_code.addItems(["PAZ", "PBZ", "PCZ", "PEZ", "PFB", "PGZ", "PHZ", "PFI", "PMZ", "POZ", "PCC"])
        self.combo_proj_code.setCurrentText("PBZ")  # پیش‌فرض
        self.combo_proj_code.setToolTip("Select or enter your project code")
        project_layout.addWidget(self.labels["proj_code"], 0, 0)
        project_layout.addWidget(self.combo_proj_code, 0, 1, 1, 2)

        # User
        self.labels["user"] = QLabel("👤 User:")
        self.line_edits["user"] = QLineEdit("PASHAEI")
        self.line_edits["user"].setToolTip("AVEVA username")
        project_layout.addWidget(self.labels["user"], 1, 0)
        project_layout.addWidget(self.line_edits["user"], 1, 1, 1, 2)

        # Password
        self.labels["password"] = QLabel("🔑 Password:")
        self.line_edits["password"] = QLineEdit("PA02")
        self.line_edits["password"].setEchoMode(QLineEdit.EchoMode.Password)
        self.line_edits["password"].setToolTip("AVEVA password")
        project_layout.addWidget(self.labels["password"], 2, 0)
        project_layout.addWidget(self.line_edits["password"], 2, 1, 1, 2)

        # MDB (فقط نمایش، خودکار پر میشه)
        self.labels["mdb"] = QLabel("💾 MDB:")
        self.line_edits["mdb"] = QLineEdit("/P2-ALL-PLANT")
        self.line_edits["mdb"].setToolTip("Master Database name (auto-filled based on project)")
        project_layout.addWidget(self.labels["mdb"], 3, 0)
        project_layout.addWidget(self.line_edits["mdb"], 3, 1, 1, 2)

        project_group.setLayout(project_layout)
        main_layout.addWidget(project_group)

        # --- Export Settings Group ---
        export_group = QGroupBox("Export Settings")
        export_group.setStyleSheet("QGroupBox { font-weight: bold; padding-top: 10px; }")
        export_layout = QGridLayout()
        export_layout.setHorizontalSpacing(15)
        export_layout.setVerticalSpacing(10)

        # Output Folder
        self.labels["output_folder"] = QLabel("📤 Output Folder:")
        self.line_edits["output_folder"] = QLineEdit("C:/Users/h.izadi/Downloads")
        self.line_edits["output_folder"].setToolTip("Folder where exported files will be saved")
        self.buttons["browse_output"] = QPushButton("📂 Browse")
        self.buttons["browse_output"].setObjectName("browse")
        export_layout.addWidget(self.labels["output_folder"], 0, 0)
        export_layout.addWidget(self.line_edits["output_folder"], 0, 1)
        export_layout.addWidget(self.buttons["browse_output"], 0, 2)

        # Navisworks Folder
        self.labels["navis_folder"] = QLabel("🔧 Navisworks Folder:")
        self.line_edits["navis_folder"] = QLineEdit("C:/Program Files/Autodesk/Navisworks Manage 2020")
        self.line_edits["navis_folder"].setToolTip("Path to Navisworks installation")
        self.buttons["browse_navis"] = QPushButton("📂 Browse")
        self.buttons["browse_navis"].setObjectName("browse")
        export_layout.addWidget(self.labels["navis_folder"], 1, 0)
        export_layout.addWidget(self.line_edits["navis_folder"], 1, 1)
        export_layout.addWidget(self.buttons["browse_navis"], 1, 2)

        # Object List
        self.labels["object_list"] = QLabel("📋 Object List (.txt):")
        self.line_edits["object_list"] = QLineEdit("")
        self.line_edits["object_list"].setToolTip("Text file containing objects to export")
        self.buttons["browse_objects"] = QPushButton("📂 Browse")
        self.buttons["browse_objects"].setObjectName("browse")
        export_layout.addWidget(self.labels["object_list"], 2, 0)
        export_layout.addWidget(self.line_edits["object_list"], 2, 1)
        export_layout.addWidget(self.buttons["browse_objects"], 2, 2)

        export_group.setLayout(export_layout)
        main_layout.addWidget(export_group)

        # --- Export Options Group ---
        options_group = QGroupBox("Export Options")
        options_group.setStyleSheet("QGroupBox { font-weight: bold; padding-top: 10px; }")
        options_layout = QVBoxLayout()
        options_layout.setSpacing(12)

        # Export Attributes Checkbox
        self.checkbox_export_attr = QCheckBox("📊 Export Attributes to Navisworks")
        self.checkbox_export_attr.setToolTip("Include element attributes in the export")
        self.checkbox_export_attr.setChecked(True)
        options_layout.addWidget(self.checkbox_export_attr)

        # Daily Export Checkbox
        self.checkbox_daily_export = QCheckBox("📅 Enable Daily Export Scheduling")
        self.checkbox_daily_export.setToolTip("Schedule automatic daily exports")
        self.checkbox_daily_export.setChecked(False)
        options_layout.addWidget(self.checkbox_daily_export)

        # Time Selection (initially hidden)
        time_layout = QHBoxLayout()
        time_layout.setContentsMargins(30, 5, 0, 5)  # Indent from left

        self.label_export_time = QLabel("⏰ Export Time:")
        self.label_export_time.setEnabled(False)

        self.time_edit_export = QTimeEdit()
        self.time_edit_export.setDisplayFormat("HH:mm")
        self.time_edit_export.setTime(QTime(0, 0))  # Default: 00:00
        self.time_edit_export.setEnabled(False)
        self.time_edit_export.setToolTip("Select the time for daily export (24-hour format)")
        self.time_edit_export.setMinimumWidth(100)

        time_layout.addWidget(self.label_export_time)
        time_layout.addWidget(self.time_edit_export)
        time_layout.addStretch()

        options_layout.addLayout(time_layout)
        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)

        # --- Action Buttons ---
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.buttons["generate"] = QPushButton("⚡ Generate & Run")
        self.buttons["generate"].setToolTip("Generate export files and execute")
        self.buttons["generate"].setMinimumWidth(150)

        self.buttons["info"] = QPushButton("ℹ️ About")
        self.buttons["info"].setToolTip("About this application")
        self.buttons["info"].setMinimumWidth(100)

        button_layout.addWidget(self.buttons["generate"])
        button_layout.addWidget(self.buttons["info"])
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        # --- Status Bar ---
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready to export | 3D DESIGN MAGIC")

        # --- Side Panel Setup ---
        self.side_panel = SidePanel(self)
        self.side_panel.setGeometry(700, 0, 300, 800)  # شروع از خارج پنجره
        self.side_panel.hide()

        # Toggle button for side panel
        self.btn_toggle_panel = QPushButton("◄", self)
        self.btn_toggle_panel.setGeometry(670, 10, 25, 50)
        self.btn_toggle_panel.setToolTip("Show/Hide Object List Panel")
        self.btn_toggle_panel.clicked.connect(self.toggle_side_panel)
        self.panel_visible = False
    def apply_dark_theme(self):
        """اعمال تم تاریک"""
        dark_stylesheet = """
        QMainWindow {
            background-color: #1E1E2E;
        }
        QWidget {
            background-color: #1E1E2E;
            color: #CDD6F4;
            font-family: 'Segoe UI', Arial;
            font-size: 10pt;
        }
        QLabel {
            color: #CDD6F4;
            font-weight: 500;
        }
        QLineEdit {
            background-color: #313244;
            border: 2px solid #45475A;
            border-radius: 6px;
            padding: 8px;
            color: #CDD6F4;
            selection-background-color: #89B4FA;
        }
        QLineEdit:focus {
            border: 2px solid #89B4FA;
        }
        QComboBox {
            background-color: #313244;
            border: 2px solid #45475A;
            border-radius: 6px;
            padding: 8px;
            color: #CDD6F4;
        }
        QComboBox:hover {
            border: 2px solid #89B4FA;
        }
        QComboBox::drop-down {
            border: none;
            width: 30px;
        }
        QComboBox::down-arrow {
            image: none;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid #CDD6F4;
            margin-right: 10px;
        }
        QComboBox QAbstractItemView {
            background-color: #313244;
            border: 2px solid #45475A;
            selection-background-color: #89B4FA;
            color: #CDD6F4;
        }
        QPushButton {
            background-color: #89B4FA;
            color: #1E1E2E;
            border: none;
            border-radius: 6px;
            padding: 10px 20px;
            font-weight: bold;
            min-width: 80px;
        }
        QPushButton:hover {
            background-color: #74C7EC;
        }
        QPushButton:pressed {
            background-color: #89DCEB;
        }
        QPushButton#browse {
            background-color: #585B70;
            color: #CDD6F4;
            min-width: 60px;
            padding: 8px 15px;
        }
        QPushButton#browse:hover {
            background-color: #6C7086;
        }
        QCheckBox {
            color: #CDD6F4;
            spacing: 8px;
        }
        QCheckBox::indicator {
            width: 20px;
            height: 20px;
            border: 2px solid #45475A;
            border-radius: 4px;
            background-color: #313244;
        }
        QCheckBox::indicator:checked {
            background-color: #89B4FA;
            border-color: #89B4FA;
        }
        QCheckBox::indicator:checked::after {
            content: "✓";
            color: #1E1E2E;
        }
                QTimeEdit {
            background-color: #313244;
            border: 2px solid #45475A;
            border-radius: 6px;
            padding: 8px;
            color: #CDD6F4;
        }
                QComboBox {
            background-color: #313244;
            border: 2px solid #45475A;
            border-radius: 6px;
            padding: 8px;
            color: #CDD6F4;
        }
        QComboBox:hover {
            border: 2px solid #89B4FA;
        }
        QComboBox:focus {
            border: 2px solid #89B4FA;
        }
        QComboBox::drop-down {
            border: none;
            padding-right: 10px;
        }
        QComboBox::down-arrow {
            image: none;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid #CDD6F4;
            margin-right: 5px;
        }
        QComboBox QAbstractItemView {
            background-color: #313244;
            border: 2px solid #45475A;
            selection-background-color: #89B4FA;
            selection-color: #1E1E2E;
            color: #CDD6F4;
        }

        QTimeEdit:focus {
            border: 2px solid #89B4FA;
        }
        QTimeEdit::up-button, QTimeEdit::down-button {
            background-color: #585B70;
            border-radius: 3px;
            width: 16px;
        }
        QTimeEdit::up-button:hover, QTimeEdit::down-button:hover {
            background-color: #6C7086;
        }
        QTimeEdit:disabled {
            background-color: #1E1E2E;
            color: #6C7086;
        }

        QMessageBox {
            background-color: #1E1E2E;
        }
        QMessageBox QLabel {
            color: #CDD6F4;
        }
        QMessageBox QPushButton {
            min-width: 70px;
        }
        """
        self.setStyleSheet(dark_stylesheet)

    def apply_light_theme(self):
        """اعمال تم روشن"""
        light_stylesheet = """
        QMainWindow {
            background-color: #FFFFFF;
        }
        QWidget {
            background-color: #FFFFFF;
            color: #2E3440;
            font-family: 'Segoe UI', Arial;
            font-size: 10pt;
        }
        QLabel {
            color: #2E3440;
            font-weight: 500;
        }
        QLineEdit {
            background-color: #FFFFFF;
            border: 2px solid #D8DEE9;
            border-radius: 6px;
            padding: 8px;
            color: #2E3440;
            selection-background-color: #5E81AC;
        }
        QLineEdit:focus {
            border: 2px solid #5E81AC;
        }
        QComboBox {
            background-color: #FFFFFF;
            border: 2px solid #D8DEE9;
            border-radius: 6px;
            padding: 8px;
            color: #2E3440;
        }
        QComboBox:hover {
            border: 2px solid #5E81AC;
        }
        QComboBox::drop-down {
            border: none;
            width: 30px;
        }
        QComboBox::down-arrow {
            image: none;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid #2E3440;
            margin-right: 10px;
        }
        QComboBox QAbstractItemView {
            background-color: #FFFFFF;
            border: 2px solid #D8DEE9;
            selection-background-color: #5E81AC;
            selection-color: #FFFFFF;
            color: #2E3440;
        }
        QPushButton {
            background-color: #5E81AC;
            color: #FFFFFF;
            border: none;
            border-radius: 6px;
            padding: 10px 20px;
            font-weight: bold;
            min-width: 80px;
        }
        QPushButton:hover {
            background-color: #81A1C1;
        }
        QPushButton:pressed {
            background-color: #4C566A;
        }
        QPushButton#browse {
            background-color: #E5E9F0;
            color: #2E3440;
            border: 1px solid #D8DEE9;
            min-width: 60px;
            padding: 8px 15px;
        }
        QPushButton#browse:hover {
            background-color: #D8DEE9;
        }
        QCheckBox {
            color: #2E3440;
            spacing: 8px;
        }
        QCheckBox::indicator {
            width: 20px;
            height: 20px;
            border: 2px solid #D8DEE9;
            border-radius: 4px;
            background-color: #FFFFFF;
        }
        QCheckBox::indicator:checked {
            background-color: #5E81AC;
            border-color: #5E81AC;
        }
        QCheckBox::indicator:checked::after {
            content: "✓";
            color: #FFFFFF;
        }
                QTimeEdit {
            background-color: #FFFFFF;
            border: 2px solid #D8DEE9;
            border-radius: 6px;
            padding: 8px;
            color: #2E3440;
        }
                QComboBox {
            background-color: #FFFFFF;
            border: 2px solid #D8DEE9;
            border-radius: 6px;
            padding: 8px;
            color: #2E3440;
        }
        QComboBox:hover {
            border: 2px solid #5E81AC;
        }
        QComboBox:focus {
            border: 2px solid #5E81AC;
        }
        QComboBox::drop-down {
            border: none;
            padding-right: 10px;
        }
        QComboBox::down-arrow {
            image: none;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid #2E3440;
            margin-right: 5px;
        }
        QComboBox QAbstractItemView {
            background-color: #FFFFFF;
            border: 2px solid #D8DEE9;
            selection-background-color: #5E81AC;
            selection-color: #FFFFFF;
            color: #2E3440;
        }

        QTimeEdit:focus {
            border: 2px solid #5E81AC;
        }
        QTimeEdit::up-button, QTimeEdit::down-button {
            background-color: #E5E9F0;
            border-radius: 3px;
            width: 16px;
        }
        QTimeEdit::up-button:hover, QTimeEdit::down-button:hover {
            background-color: #D8DEE9;
        }
        QTimeEdit:disabled {
            background-color: #F5F5F5;
            color: #A0A0A0;
        }

        QMessageBox {
            background-color: #FFFFFF;
        }
        QMessageBox QLabel {
            color: #2E3440;
        }
        QMessageBox QPushButton {
            min-width: 70px;
        }
        """
        self.setStyleSheet(light_stylesheet)

    def connect_signals(self):
        """
        Connects widget signals to corresponding slots (functions).
        """
        # Theme change
        self.combo_mode.currentTextChanged.connect(self.on_theme_changed)

        # Project Code auto-fill MDB
        self.combo_proj_code.currentTextChanged.connect(self._on_project_changed)

        # Browse buttons
        self.buttons["browse_aveva"].clicked.connect(lambda: self._browse_folder(self.line_edits["aveva_folder"]))
        self.buttons["browse_output"].clicked.connect(lambda: self._browse_folder(self.line_edits["output_folder"]))
        self.buttons["browse_navis"].clicked.connect(lambda: self._browse_folder(self.line_edits["navis_folder"]))
        self.buttons["browse_objects"].clicked.connect(
            lambda: self._browse_file(self.line_edits["object_list"], "Text files (*.txt)"))

        # Main action buttons
        self.buttons["generate"].clicked.connect(self.generate_files)
        self.buttons["info"].clicked.connect(self._show_info_dialog)

        # Daily export checkbox controls time edit visibility
        self.checkbox_daily_export.stateChanged.connect(self._on_daily_export_changed)

        # Project code change loads corresponding object list
        self.combo_proj_code.currentTextChanged.connect(self.load_panel_objects)

        # Save panel objects when list changes
        self.side_panel.list_widget.model().rowsInserted.connect(self.save_panel_objects)
        self.side_panel.list_widget.model().rowsRemoved.connect(self.save_panel_objects)

    def _on_project_changed(self, project_code):
        """Auto-fill MDB based on selected project code."""
        project_mdb_mapping = {
            "PAZ": "/P1-ALL-PLANT",
            "PBZ": "/P2-ALL-PLANT",
            "PCZ": "/P3-ALL-PLANT",
            "PEZ": "/P5-ALL-PLANT-SU",
            "PFB": "/ALL-PLANT-P6.2-NEW",
            "PGZ": "/P7-ALL-PLANT",
            "PHZ": "/P8-ALL-PLANT",
            "PFI": "/P04-ALL",
            "PMZ": "/P12-ALL-PLANT-KHARG",
            "POZ": "/ALL-PLANT-P13",
            "PCC": "/ALL-2140"
        }

        # اگر پروژه در mapping باشد، MDB رو پر کن
        if project_code in project_mdb_mapping:
            self.line_edits["mdb"].setText(project_mdb_mapping[project_code])

    def load_theme(self):
        """بارگذاری و اعمال تم ذخیره شده"""
        saved_theme = self.settings.value('theme', 'Dark')  # پیش‌فرض Dark
        self.combo_mode.setCurrentText(saved_theme)
        self.on_theme_changed(saved_theme)

        # Load daily export settings
        daily_export_enabled = self.settings.value('daily_export', False, type=bool)
        self.checkbox_daily_export.setChecked(daily_export_enabled)

        # Load saved time
        saved_time = self.settings.value('export_time', '00:00', type=str)
        time_obj = QTime.fromString(saved_time, "HH:mm")
        if time_obj.isValid():
            self.time_edit_export.setTime(time_obj)

        # Update time edit state based on checkbox
        self.time_edit_export.setEnabled(daily_export_enabled)
        self.label_export_time.setEnabled(daily_export_enabled)

        # Load panel objects for current project
        self.load_panel_objects()

    def on_theme_changed(self, theme_name):
        """تغییر تم بر اساس انتخاب کاربر"""
        if theme_name == "Light":
            self.apply_light_theme()
        else:
            self.apply_dark_theme()

        # ذخیره تنظیمات
        self.settings.setValue('theme', theme_name)

    def _browse_folder(self, line_edit):
        """Opens a dialog to select a folder and sets the path to the line edit."""
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            line_edit.setText(folder_path.replace("\\", "/"))

    def _browse_file(self, line_edit, file_filter):
        """Opens a dialog to select a file and sets the path to the line edit."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", file_filter)
        if file_path:
            line_edit.setText(file_path.replace("\\", "/"))

    def _show_info_dialog(self):
        """Displays a professional About dialog with detailed information."""
        dialog = QDialog(self)
        dialog.setWindowTitle("About - Export E3D to Navisworks")
        dialog.setFixedSize(500, 450)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Logo/Header
        header = QLabel("🏗️ Export E3D to Navisworks")
        header_font = QFont()
        header_font.setPointSize(16)
        header_font.setBold(True)
        header.setFont(header_font)
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header)

        # Version
        version_label = QLabel("Version 1.0")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_font = QFont()
        version_font.setPointSize(10)
        version_label.setFont(version_font)
        layout.addWidget(version_label)

        # Separator
        separator = QLabel("━" * 60)
        separator.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(separator)

        # Description
        description = QTextBrowser()
        description.setMaximumHeight(150)
        description.setOpenExternalLinks(False)
        description_text = """
        <p><b>پلتفرم حرفه‌ای اکسپورت AVEVA E3D به Navisworks</b></p>

        <p>این ابزار قدرتمند برای تبدیل خودکار مدل‌های AVEVA E3D به فرمت Navisworks طراحی شده است. 
        با قابلیت‌های پیشرفته شامل:</p>

        <ul>
            <li>تولید خودکار فایل‌های RVM و Attribute</li>
            <li>پشتیبانی از Export دسته‌جمعی اشیاء</li>
            <li>ذخیره خودکار تنظیمات</li>
        </ul>
        """
        description.setHtml(description_text)
        layout.addWidget(description)

        # Developer Info
        info_text = QLabel()
        info_text.setWordWrap(True)
        info_text.setTextFormat(Qt.TextFormat.RichText)
        info_text.setOpenExternalLinks(False)
        info_text.setText("""
        <p><b>👨‍💻 توسعه‌دهنده:</b> HOSSEIN IZADI</p>
        <p><b>📧 ایمیل:</b> <a href='hossein.pi.ad@gmail.com' style='color: #5E81AC; text-decoration: none;'>hossein.pi.ad@gmail.com</a></p>
        <p><b>🌐 گیت‌هاب:</b> <a href='https://github.com/arkittioe' style='color: #5E81AC; text-decoration: none;'>https://github.com/arkittioe</a></p>
        <p><b>📅 تاریخ انتشار:</b> 1404/08/24</p>
        """)

        # Connect link clicks
        info_text.linkActivated.connect(self._open_link)
        layout.addWidget(info_text)

        # Footer
        footer = QLabel("Made with ❤️ for AVEVA & Navisworks users")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_font = QFont()
        footer_font.setItalic(True)
        footer_font.setPointSize(9)
        footer.setFont(footer_font)
        layout.addWidget(footer)

        layout.addStretch()

        # Close Button
        close_button = QPushButton("✓ Close")
        close_button.clicked.connect(dialog.accept)
        close_button.setMinimumWidth(100)
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(close_button)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        dialog.exec()

    def _open_link(self, url):
        """Opens URLs in default browser."""
        QDesktopServices.openUrl(QUrl(url))

    def generate_files(self):
        """
        Main logic to generate the four output files based on GUI inputs.
        """
        self.status_bar.showMessage("⏳ Generating files...")
        try:
            # --- 1. Collect and Validate Data ---

            # Get export time based on checkbox state
            if self.checkbox_daily_export.isChecked():
                export_time_value = self.time_edit_export.time().toString("HH:mm")
            else:
                export_time_value = False

            data = {
                "aveva_path": self.line_edits["aveva_folder"].text().strip(),
                "proj_code": self.combo_proj_code.currentText().strip(),
                "user": self.line_edits["user"].text().strip(),
                "password": self.line_edits["password"].text(),
                "mdb": self.line_edits["mdb"].text().strip(),
                "output_folder": self.line_edits["output_folder"].text().strip(),
                "roamer_path": os.path.join(self.line_edits["navis_folder"].text().strip(), "Roamer.exe").replace("\\",
                                                                                                                  "/"),
                "areas_file": self.line_edits["object_list"].text().strip(),
                "export_attribute": self.checkbox_export_attr.isChecked(),
                "daily_export": self.checkbox_daily_export.isChecked(),
                "export_time": export_time_value
            }

            # Basic validation (skip export_time and areas_file)
            for key, value in data.items():
                if key in ["export_time", "areas_file"]:  # Skip optional fields
                    continue
                if not value and isinstance(value, str):
                    QMessageBox.warning(self, "Input Error",
                                        f"Field '{key.replace('_', ' ').title()}' cannot be empty.")
                    return

            output_folder = Path(data["output_folder"])
            if not output_folder.exists() or not output_folder.is_dir():
                QMessageBox.warning(self, "Path Error", f"Output folder does not exist:\n{output_folder}")
                return

            # --- 2. Read Object List (from file OR panel) ---
            object_list = []
            areas_file_path = Path(data["areas_file"]) if data["areas_file"] else None

            # اول سعی کن از فایل بخون
            if areas_file_path and areas_file_path.exists() and areas_file_path.is_file():
                with open(areas_file_path, 'r', encoding='utf-8') as f:
                    object_list = [line.strip() for line in f if line.strip()]

            # اگر فایل نبود یا خالی بود، از پنل استفاده کن
            if not object_list:
                object_list = self.side_panel.get_items()

            # چک نهایی: اگه هیچ object نداریم
            if not object_list:
                QMessageBox.warning(
                    self,
                    "Object List Error",
                    "No objects found!\n\n"
                    "Please either:\n"
                    "1. Specify a valid SITE.txt file with objects, or\n"
                    "2. Add objects to the side panel (click ◄ button)"
                )
                return

            # Normalize paths for macro files (use forward slashes)
            normalized_data = {k: (v.replace("\\", "/") if isinstance(v, str) else v) for k, v in data.items()}

            # --- 3. Generate Files ---
            self._generate_settings_json(output_folder, normalized_data)
            self._generate_rvm_mac(output_folder, normalized_data, object_list)
            self._generate_attribute_mac(output_folder, normalized_data, object_list)
            self._generate_run_bat(output_folder, normalized_data)

            QMessageBox.information(self, "Success", f"All files were generated successfully in:\n{output_folder}")
            self.status_bar.showMessage("✓ Files generated successfully!", 5000)

            # --- 4. اجرای خودکار فایل RunE3D.bat ---
            bat_file_path = output_folder / "RunE3D.bat"
            if bat_file_path.exists():
                try:
                    # استفاده از os.startfile برای اجرای فایل به صورت مستقل (مثل دوبار کلیک)
                    os.startfile(str(bat_file_path))
                    self.status_bar.showMessage("✓ RunE3D.bat executed successfully!", 3000)
                except Exception as e:
                    QMessageBox.warning(self, "Execution Error",
                                        f"Files generated but failed to run RunE3D.bat:\n{str(e)}")
            else:
                QMessageBox.warning(self, "File Error",
                                    "RunE3D.bat was not generated properly.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{str(e)}")
            self.status_bar.showMessage("✗ Error occurred during generation", 5000)

    def _generate_settings_json(self, output_dir, data):
        """Generates the settings.json file."""
        settings_path = output_dir / "settings.json"
        # Create a dictionary with only the keys needed for the JSON file
        json_data = {
            "aveva_path": data["aveva_path"],
            "proj_code": data["proj_code"],
            "user": data["user"],
            "password": data["password"],
            "mdb": data["mdb"],
            "output_folder": data["output_folder"],
            "roamer_path": data["roamer_path"],
            "areas_file": data["areas_file"],
            "export_attribute": data["export_attribute"],
            "daily_export": data["daily_export"],
            "export_time": data["export_time"],
        }
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=4)

    def _generate_rvm_mac(self, output_dir, data, objects):
        """Generates the RVM.mac file."""
        log_file_path = f"{data['output_folder']}/RVM_LOG.txt".replace("/", "\\")
        attribute_mac_path = f"{data['output_folder']}/attribute.mac"
        temp_rvm_path = f"{data['output_folder']}/TEMP.RVM"

        # Dynamically create the EXPORT commands
        export_commands = "\n".join(
            f"SYSCOM |echo [RVM] Export {obj} >> {log_file_path}|\nEXPORT {obj}" for obj in objects
        )

        content = f"""
DESIGN

SYSCOM |echo [RVM] Start attribute.mac >> {log_file_path}|
$M {attribute_mac_path}

VAR !PROJ PROJ CODE
!CUDATE = OBJECT DATETIME()
!DAY = !CUDATE.DATE().STRING()
IF !DAY.LENGTH().EQ( 1 ) THEN
  !DAY = '0' + !DAY
ENDIF
!MONTH = !CUDATE.MONTH().STRING()
IF !MONTH.LENGTH().EQ( 1 ) THEN
  !MONTH = '0' + !MONTH
ENDIF

!FILNAME = '{data['output_folder']}/' + '$!PROJ-' + !CUDATE.YEAR().STRING()+ '-' + !MONTH + '-' + !DAY + '.nwd'
SYSCOM |echo [RVM] NWD_OUT=$!FILNAME >> {log_file_path}|

SYSCOM |echo [RVM] Exporting RVM file... >> {log_file_path}|
EXPORT FILE /{temp_rvm_path} OVER
EXPORT AUTOCOLOUR DISPLAYEXPORT ON
EXPORT AUTOCOLOUR ON
EXPORT REPR ON
EXPORT HOLES ON
EXPORT IMPLIED TUBE INTO SEPARATE
{export_commands}
EXPORT FINISH

SYSCOM |echo [RVM] Launching Navisworks... >> {log_file_path}|
SYSCOM |""{data['roamer_path']}" -nwd $!FILNAME "{temp_rvm_path}"|

SYSCOM |echo [RVM] Finished >> {log_file_path}|
FINISH
"""
        with open(output_dir / "RVM.mac", 'w', encoding='utf-8') as f:
            f.write(content)

    def _generate_attribute_mac(self, output_dir, data, objects):
        """Generates the attribute.mac file."""
        temp_txt_path = f"{data['output_folder']}/TEMP.txt"

        # Join the object list into a space-separated string for the 'collect all' command
        objects_string = " ".join(objects)

        content = f"""
onerror continue
$:debug$:
$* modificato 12-04-2022
$* Initialise Variables
var !FILE |{temp_txt_path}|
var !DILM |:=|
var !SEPR |&end&|
var !IGNORE |,NAME,OWNER,|
var !AIGNORE |,unset,=0/0,nulref,|
var !ODEPTH DDEPTH
var !PDEPTH -99
var !REFE REFE
var !count 1

$* Open file
openfile /$!FILE write !FUNIT
handle ANY
openfile /$!FILE overwrite !FUNIT
endhandle

$* Write Header
var !DATE clock date
var !TIME clock time
writefile $!FUNIT |CADC_Attributes_File v1.0 , start: NEW , end: END , name_end: $!DILM , sep: $!SEPR|
writefile $!FUNIT |NEW Header Information|

var !DPRT compose | Source$!DILM PDMS Data $!SEPR Date$!DILM $!DATE$n $!SEPR Time$!DILM $!TIME|
writefile $!FUNIT |$!DPRT[1]|
var !MDB MDB
var !PROJECT PROJECT CODE
var !NAME (FULLNAME)
var !DPRT compose | Project$!DILM $!PROJECT $!SEPR MDB$!DILM $!MDB $!SEPR |
writefile $!FUNIT |$!DPRT[1]|
writefile $!FUNIT |END|

!list = 'SITE ZONE PIPE BRAN ELBOW BEND TEE FLAN OLET INST VALVE PCOMP FBLIND GASK TUBI REDU CAP COUP PLUG UNION ATTA FTUBE FILT STRU FRMW SCTN'

$* Get element
Var !COLL collect all ($!list) for {objects_string}

$* Loop through the list of elements
do !INDX indices !COLL

$!COLL[$!INDX]

$* Hierarchy level
var !DEPTH DDEPTH
var !TAB $!DEPTH * 2 - $!ODEPTH * 2
var !ITAB $!TAB + 2

$* End(s)
if($!DEPTH eq $!PDEPTH and $!PDEPTH neq -99) then
var !DPRT compose space $!TAB |END|
writefile $!FUNIT |$!DPRT[1]|
elseif ($!DEPTH lt $!PDEPTH) then

do !INDXA from $!PDEPTH to $!DEPTH by -1
var !DPRT compose space $!INDXA |END|
writefile $!FUNIT |$!DPRT[1]|
enddo
endif

var !PDEPTH $!DEPTH

$* Attributes of element  (this is new)
var !ATTL delete
IF (TYPE eq |TUBI|) THEN
var !ATTL append |Itlength|
var !ATTL append |Lbore|
var !ATTL append |DTXR|
var !ATTL append |Spref|
ELSE
var !ATTL attlist
ENDIF

-- initialise progress and interrupt system
!progress = 0
!progStep = 5 $* % progress report step
$*!this.enableInterrupt()

$*onerror golabel /interrupted
$* Check it item is owned by a branch
if(TYPE neq |WORL|) then
if(TYPE of OWNER eq |BRAN| and NOT BADREF(SPREF)) then
if (TYPE neq |TUBI|) then
var !ATTL append |APOS|
var !ATTL append |LPOS|
var !ATTL append |DTXR|
var !ATTL append |CWEI OF CMPREF OF SPREF|
var !ATTL append |AFTER(NAME OF PSPEC OF PIPE, '/')|
$* var !ATTL append |:PNUM of spco of spref|
$* var !ATTL append |:ENI_CODE of spco of spref|
var !ATTL append |P1BOR|
var !ATTL append |P2BOR|
if(type eq |TEE|) or (type eq|OLET|)then
var !ATTL append |P3BOR|
endif
endif
endif
endif

$* Get name
var !NAME (FULLNAME)
if (TYPE eq |TUBI|) then
var !nametube coll all tubi for owne
!countTubi = !nametube.FindFirst(!COLL[$!INDX])
var !NAME NAME OF BRANCH
var !DPRT compose space $!TAB |NEW TUBE $!countTubi of BRANCH $!NAME|
else
var !DPRT compose space $!TAB |NEW $!NAME|
endif
writefile $!FUNIT |$!DPRT[1]|
var !ASIZE (arraywidth(!ATTL)) + 3

$* Loop through attribute array
do !ATTR values !ATTL

skip if(match(|$!IGNORE|,|,$!ATTR$n,|) gt 0)
var !ATTRIB (ATTRIB $!ATTR)
handle ANY

var !ATTRIB $!ATTR
$* replased text to have the best readible in Navis (this is new)
endhandle
var !ATTRIB (trim(|$!ATTRIB|))
if(|$!ATTRIB| neq || and match(|$!AIGNORE|,|,$!ATTRIB$n,|) eq 0) then
var !new REPLACE (|$!ATTR$n|,|Itlength|,|Length|)
var !new REPLACE (|$!new$n|,|CWEI OF CMPREF OF SPREF|,|weight|)
var !new REPLACE (|$!new$n|,|DTXR|,|Descr.|)
var !new REPLACE (|$!new$n|,|P2BOR|,|Red. Size|)
var !new REPLACE (|$!new$n|,|P3BOR|,|Branch Conn. Size|)
var !new REPLACE (|$!new$n|,|AFTER(NAME OF PSPEC OF PIPE, '/')|,|Pipe Spec|)
var !new REPLACE (|$!new$n|,|:ENI_CODE of spco of spref|,|ENI Code|)
var !new REPLACE (|$!new$n|,|:PNUM of spco of spref|,|PUMA Code|)
var !new REPLACE (|$!new$n|,|P1BOR|,|Main Size|)
var !new REPLACE (|$!new$n|,|Lbore|,|Pipe Size|)
if(|$!new$n| eq |Length|)then
var !ATTRIB $!ATTRIB
var !ATTRIB STRING ( $!ATTRIB, 'D2' )
var !ATTRIB |$!ATTRIB mm.|
endif

var !DPRT compose space $!ITAB |$!new$!DILM| width $!ASIZE R space 2 |$!ATTRIB|

!size = !COLL.size()
-- Update progress if required
!percentDone = int( (!INDX * 100 ) / $!size )
--$P index = $!indx, %done is $!percentDone, step $!progStep%
if( !percentDone - !progress ge !progStep ) then
!progress = !percentDone
--$P $!progress%
!!fmsys.setProgress( !progress )
endif

writefile $!FUNIT |$!DPRT[1]|
endif
enddo

enddo

if($!DEPTH gt $!ODEPTH) then
do !INDXA from $!DEPTH to $!ODEPTH by -1
var !TAB $!INDXA * 2 - $!ODEPTH * 2
var !DPRT compose space $!TAB |END|
writefile $!FUNIT |$!DPRT[1]|
enddo
endif

$!REFE

closefile $!FUNIT

$* Hide the from
$* hide _CDXATTDUMP
!!fmsys.setProgress( 0 )
$P finish
return $* >>>>>>>>>> End of Code DesignReview <<<<<<<<<<
$.
"""
        with open(output_dir / "attribute.mac", 'w', encoding='utf-8') as f:
            f.write(content.strip())

    def _generate_run_bat(self, output_dir, data):
        """
        Generates the RunE3D.bat file using the user-provided advanced template.
        This version tracks macro progress and waits for the NWD file.
        After successful completion, it deletes all generated files permanently (except RVM_LOG.txt).
        """
        # Ensure paths are correctly formatted for the batch script (using backslashes)
        output_folder_bat = data["output_folder"].replace("/", "\\")
        aveva_path_bat = data["aveva_path"].replace("/", "\\")

        log_file_path = os.path.join(output_folder_bat, "RVM_LOG.txt")
        monitor_path = os.path.join(aveva_path_bat, "mon.exe")
        launch_init_path = os.path.join(aveva_path_bat, "launch.init")
        rvm_mac_path = os.path.join(output_folder_bat, "RVM.mac")

        # The user's credentials and project info
        proj_code = data["proj_code"]
        user = data["user"]
        password = data["password"]
        mdb = data["mdb"]

        # Paths for the files to be deleted after completion (in reverse order)
        settings_json_path = os.path.join(output_folder_bat, "settings.json")
        rvm_mac_delete_path = rvm_mac_path
        attribute_mac_path = os.path.join(output_folder_bat, "attribute.mac")
        temp_txt_path = os.path.join(output_folder_bat, "TEMP.txt")
        temp_rvm_path = os.path.join(output_folder_bat, "TEMP.RVM")
        bat_file_path = os.path.join(output_folder_bat, "RunE3D.bat")

        # This content is the user's template, with dynamic values injected.
        content = f"""
    @echo off
    setlocal ENABLEDELAYEDEXPANSION
    echo [INFO] Starting AVEVA E3D with macro RVM.mac...
    echo -----------------------------------------

    if exist "{log_file_path}" del "{log_file_path}"

    start "" /b "{monitor_path}" ^
          PROD E3D init "{launch_init_path}" ^
          GRAPHICS {proj_code} {user}/{password} /{mdb} ^
          $M{rvm_mac_path}

    echo [INFO] Tracking macro progress...
    set "NWD_PATH="
    :loop
    if exist "{log_file_path}" (
      rem The 'type' command can sometimes lock the file, so we'll be careful.
      rem Instead of constantly typing, we just check for the final string.
      findstr /c:"[RVM] Finished" "{log_file_path}" >nul
      if not errorlevel 1 (
         echo [INFO] Macro has finished. Checking for NWD file path...
         for /f "usebackq tokens=1,* delims==" %%A in (`findstr /c:"[RVM] NWD_OUT=" "{log_file_path}"`) do (
            set "NWD_PATH=%%B"
         )
         if defined NWD_PATH (
            echo [INFO] NWD path found: !NWD_PATH!
            if exist "!NWD_PATH!" goto done
            echo [WARN] NWD file does not exist yet, waiting...
         )
      )
    )
    echo [INFO] Waiting for macro to complete... (checking again in 5s)
    timeout /t 5 >nul
    goto loop

    :done
    echo.
    echo [SUCCESS] Process finished. NWD file is ready at: !NWD_PATH!
    echo -----------------------------------------

    rem Wait a moment to ensure all file handles are released
    timeout /t 2 >nul

    echo [INFO] Cleaning up generated files...
    rem Delete files in reverse order: attribute.mac, RVM.mac, settings.json, TEMP.txt, TEMP.RVM
    if exist "{attribute_mac_path}" (
        del /f /q "{attribute_mac_path}"
        echo [INFO] Deleted: attribute.mac
    )
    if exist "{rvm_mac_delete_path}" (
        del /f /q "{rvm_mac_delete_path}"
        echo [INFO] Deleted: RVM.mac
    )
    if exist "{settings_json_path}" (
        del /f /q "{settings_json_path}"
        echo [INFO] Deleted: settings.json
    )
    if exist "{temp_txt_path}" (
        del /f /q "{temp_txt_path}"
        echo [INFO] Deleted: TEMP.txt
    )
    if exist "{temp_rvm_path}" (
        del /f /q "{temp_rvm_path}"
        echo [INFO] Deleted: TEMP.RVM
    )

    echo [INFO] Cleanup complete. This batch file will now self-destruct...
    rem Self-delete this batch file (RunE3D.bat)
    (goto) 2>nul & del /f /q "%~f0"
    """
        with open(output_dir / "RunE3D.bat", 'w', encoding='utf-8') as f:
            f.write(content.strip())

    def _on_daily_export_changed(self, state):
        """Enable/disable time selection based on daily export checkbox."""
        is_enabled = (state == Qt.CheckState.Checked.value)
        self.time_edit_export.setEnabled(is_enabled)
        self.label_export_time.setEnabled(is_enabled)

        # Save state to settings
        self.settings.setValue('daily_export', is_enabled)
        if is_enabled:
            self.settings.setValue('export_time', self.time_edit_export.time().toString("HH:mm"))

    def get_default_objects(self):
        """
        Returns default object lists for each project code.
        """
        default_objects = {
            "PAZ": [
                "/BOCRefSite/X45dc-SE3P",
                "/BOCRefSite/X45dc-SE3P2",
                "/PSI-SAMPLE",
                "/U106A:EL",
                "/U106A:SU",
                "/U106A:EQ",
                "/SUPP-106A",
                "/ST-SUPP",
                "/U106A:IN",
                "/U106A:SA",
                "/U106A:FF",
                "/U106:PI-SA-EDITED",
                "/U106A:ST",
                "/U106A:LP",
                "/U106A:CF",
                "T /U106A:CI"
            ],
            "PBZ": [
                "/SUPPORT-109A-PI",
                "/SUP-ST109A",
                "/SUPPORT-109A-UTILITY",
                "/FOUNDATION-SUPPORT",
                "/U109A_EQ",
                "/U109A:EL",
                "/U109A:IN",
                "/U109A:PI",
                "/U109A:SA",
                "/U109A:FF",
                "/U109A:ST",
                "/U109A:LP",
                "/U109A:CF",
                "/OVR04:PU"
            ],
            "PCZ": [
                "/U102A:EL",
                "/U102A:EQ",
                "/U102A:IN",
                "/U102A:PI",
                "/U102A:SA",
                "/U102A:FF",
                "/U102A:ST",
                "/U102A:LP",
                "/U102A:CF",
                "/U102A:CA",
                "/CP6I:Civil(P12)",
                "/OVR02:CA(P12)",
                "/OVR01:CA(P12)",
                "/U102B:PI",
                "/EQ-SUPP",
                "/U104A:EL",
                "/U104A:EQ",
                "/U104A:IN",
                "/U104A:PI",
                "/U104A:CA",
                "/U104A:SA",
                "/U104A:FF",
                "/U104A:ST",
                "/U104A:LP",
                "/U104A:CF",
                "/U107A:CA",
                "/U107A:EL",
                "/U107A:EQ",
                "/U107A:IN",
                "/U107A:PI",
                "/U107A:SA",
                "/U107A:FF",
                "/U107A:ST",
                "/U107A:LP",
                "/U107A:CF",
                "/U107B:CA",
                "/U107B:ST",
                "/U107B:LP",
                "/U108B:EL",
                "/U108B:EQ",
                "/U108B:IN",
                "/U108B:PI",
                "/U108B:SA",
                "/U108B:FF",
                "/U108B:ST",
                "/U108B:LP",
                "/U108B:CF",
                "/U108B:CA",
                "/FRAM-ST "
            ],
            "PEZ": [
                "/U102B:IN",
                "/MJ",
                "/SUPP-P05-NEW",
                "/SUPP-EQ.NEW",
                "/U102B:EL",
                "/SUPP-FOUNDATION",
                "/ST-SUPPORT-NEW",
                "/U102B:EQ",
                "/U102B:PROPANE:EQ",
                "/U102B:BUTANE:EQ",
                "/FIREWATER-PIPE",
                "/U102B:PI",
                "/U102B:SA",
                "/U102B:FF",
                "/U102B:ST",
                "/U102B:LP",
                "/U102B:CF",
                "/U102B:CU",
                "/U103B:EL",
                "/U103B:EQ",
                "/U103B:IN",
                "/U103B:PI",
                "/U103B:SA",
                "/U103B:FF",
                "/U103B:ST",
                "/U103B:LP",
                "/U103B:CF",
                "/U104A:FF",
                "/U104B:EL",
                "/U104B:EQ",
                "/U104B:IN",
                "/U104B:PI",
                "/U104B:SA",
                "/U104B:FF",
                "/U104B:ST",
                "/U104B:LP",
                "/U104B:CF",
                "/U105B:EL",
                "/U105B:EQ",
                "/U105B:IN",
                "/U105B:PI",
                "/U105B:SA",
                "/U105B:FF",
                "/U105B:ST",
                "/U105B:LP",
                "/U105B:CF",
                "/U108C:EL",
                "/U108C:EQ",
                "/U108C:IN",
                "/U108C:PI",
                "/U108C:SA",
                "/U108C:FF",
                "/U108C:ST",
                "/U108C:LP",
                "/U108C:CF"
            ],
            "PFB": [
                "/AWNING",
"/U101A:PI",
"/FARAB-PIPING",
"/U101A:SA",
"/U101A:FF",
"/BOILER-D",
"/OVERALL",
"/BOILER-C",
"/BOILER-B",
"/BOILER-A",
"/PI",
"/FARAB-CIVIL",
"/FARAB1-EQUIPMENTS",
"/FARAB2-EQUIPMENTS",
"/CA",
"/EQ",
"/ST",
"/LP",
"/CF",
"/EL",
"/IN",
"/DOSING",
"/PU",
"/SUPPORT-P06",
"/EQ-SUPPORT",
"/ST-SUPPORT",
"/T-PIPE",
"/SITE-AIR",
"/PI/RO",
"/DIESEL-FUEL-PACKAGE",
"/BOCAD",
"/MJ",
"/SITEGTG",
"/U101A:EL",
"/U101A:EQ",
"/EQUIPMENT-101A",
"/DEISEL-FILTER",
"/U101A:IN",
"/U101A:ST",
"/U101A:LP",
"/U101A:CF",
"/OVR03:CA(PFB)",
"/P12",
"/OVR03:CV(PFB)",
"/CP62:TRN",
"/101A-UG",
"EXCLUDE /12-WF-240034-D1C-UW(OVR03)(P12-INT)",
"EXCLUDE /12-WF-240037-D1C-UW(OVR03)(P12-INT)",
"EXCLUDE /PR-101A-06",
"EXCLUDE /PR-101A-05",
"EXCLUDE /TIEIN-STR"
            ],
            "PGZ": [
                "/OVRP7:CU",
"/DOSING_UG",
"/OVRP7:PU",
"/UP7:EL",
"/UP7:EQ",
"/UP7:IN",
"/UP7:PI",
"/UP7:SA",
"/UP7:FF",
"/UP7:ST",
"/UP7:LP",
"/UP7:CF",
"/AREA-2433",
"/UG-MTO-SITE",
"/AREA-2413",
"EXCLUDE /6-WF-230507-D1C-UW(101A)",
"EXCLUDE /6-WF-230508-D1C-UW(101A)"
            ],
            "PHZ": [
                "/JETTY-KHARG",
"/PFB-SUPPORT",
"/CP8:TRN",
"/UP8:FOUN",
"/UP8:ST",
"/UP8:LP",
"/UP8:CF",
"/UP8:EL",
"/UP8:IN",
"/UP8:EQ",
"/UP8:BL",
"/KHARG-P08-SUPPORT",
"/SUPPORT-P8",
"/EQ-SUPPORT",
"/UP8:PI/JETTY",
"/UP8:SA",
"/UP8:PI/SEAWATER-INTAKE",
"/UP8:PI/BOG",
"/OVRP8:CA"
            ],
            "PFI": [
                "/U101A:SA",
"/BOCAD-110A",
"/BOCRefSite/X150c-L-SAEEDI",
"/PSI-SAMPLE",
"/SDFFF",
"/BOCRefSite/X45dc-SE3P",
"/BOCRefSite/X45dc-SE3P2",
"/M.J",
"/OVR01:CA",
"/CP6I:Civil",
"/P12-ELEC.",
"/P12-ROAD",
"/PIPING-SEA-WATER",
"/OVR03:PU(PFB)",
"/OVR01:CU",
"/OVR01:PU",
"/OVR02:CA",
"/OVR02:CU",
"/OVR02:PU",
"/OVR03:CA(PFI)",
"/U102D:ST",
"/U102D:PI",
"/OVR03:CU(PFI)",
"/OVR03:PU(PFI)",
"/OVR04:CA",
"/OVR04:CU",
"/OVR04:PU",
"/OVR05:CA",
"/OVR05:CU",
"/OVR05:PU",
"/OVR06:CA",
"/OVR06:CU",
"/OVR06:PU",
"/U108A:EL",
"/U108A:ST",
"/U108D:ST",
"/U108D:PI",
"/Copy-of-OVR06:PU",
"/BOCAD",
"/OVR07:CA",
"/OVR07:CU",
"/OVR07:PU",
"/U100A:EL",
"/U100A:EQ",
"/U100A:IN",
"/U100C:PI",
"/U100A:SA",
"/U100A:FF",
"/U100A:ST",
"/U100A:LP",
"/U100A:CF",
"/U100B:EL",
"/U100B:EQ",
"/U100B:IN",
"/U100B:PI",
"/U100B:SA",
"/U100B:FF",
"/U100B:ST",
"/U100B:LP",
"/U100B:CF",
"/U100C:EL",
"/U100C:EQ",
"/U100C:IN",
"/U100A:PI",
"/U100C:SA",
"/U100C:FF",
"/U100C:ST",
"/U100C:LP",
"/U100C:CF",
"/U100D:EL",
"/U100D:EQ",
"/U100D:IN",
"/U100D:PI",
"/U100D:SA",
"/U100D:FF",
"/U100D:ST",
"/U100D:LP",
"/U100D:CF",
"/U100E:EL",
"/U100E:EQ",
"/U100E:IN",
"/U100E:PI",
"/U100E:SA",
"/U100E:FF",
"/U100E:ST",
"/U100E:LP",
"/U100E:CF",
"/U106A:CI",
"/U106B:EL",
"/U106B:EQ",
"/U106B:IN",
"/U106B:PI",
"/U106B:SA",
"/U106B:FF",
"/U106B:ST",
"/U106B:LP",
"/U106B:CF",
"/U106C:EL",
"/U106C:EQ",
"/U106C:IN",
"/U106C:PI",
"/U106C:SA",
"/U106C:FF",
"/U106C:ST",
"/U106C:LP",
"/U106C:CF",
"/U106D:EL",
"/U106D:EQ",
"/U106D:IN",
"/U106D:PI",
"/U106D:SA",
"/U106D:FF",
"/SUPPORT-P12",
"/ST-SUPORT-P12",
"/U106D:ST",
"/U106D:LP",
"/U106D:CF",
"/U106E:EL",
"/U106E:EQ",
"/U106E:IN",
"/U106E:SA",
"/U106E:FF",
"/U106E:PI",
"/U106E:ST",
"/U106E:LP",
"/U106E:CF",
"/U101E:EQ",
"/U101E:EL",
"/U101E:IN",
"/U101E:PI",
"/U101E:SA",
"/U101E:FF",
"/U101E:ST",
"/U101E:LP",
"/U101E:CF",
"/U101C:EL",
"/U101C:EQ",
"/U101C:IN",
"/U101C:PI",
"/U101C:SA",
"/U101C:FF",
"/U101C:ST",
"/U101C:LP",
"/U101C:CF",
"/U102C:PI",
"/U102C:EQ",
"/U104C:EQ",
"/U104C:PI",
"/SDNF-CONFIG-P12-BOCAD-BEHROOZI",
"/U110A:EL",
"/U110A:EQ",
"/U110A:IN",
"/U110A:PI",
"/U110A:SA",
"/U110A:FF",
"/U110A:ST",
"/U101B:ST",
"/U110A:LP",
"/U110A:CF",
"/OUTFALL",
"/BOCAD-110-2",
"EXCLUDE /6-WF-230507-D1C-UW(U101A)",
"EXCLUDE /12-WF-240021-D1C-UW(OVR05)(P12-INT)",
"EXCLUDE /12-WF-240080-D1C-UW(OVR05)(P12-INT)",
"EXCLUDE /8-WF-240074-D1C-UW(OVR05)(P12-INT)",
"EXCLUDE /6-WF-230509-D1C-UW(U101A)"
            ],
            "PMZ": [
                "/SRU-CIVIL",
"/P13-PIPING",
"/P13-EQUIPMENT",
"/SRU-STRU",
"/CN"
            ],
            "POZ": [
                "/UPOZZEQDES/EQ",
"/UPOZZSTDES/ST",
"/UPOZZELDES/EL",
"/UPOZZELDES/IN",
"/M.J",
"/UPOZZPIDES/PI"
            ],
            "PCC": [
                "/UG-PIP",
"/KHORAVI-A",
"/CP6I-Civil",
"/SITE-2140",
"/PIPING-2139",
"/EQUIPMENTS-2139",
"/STRUTCURES-2139",
"/PLATFORM-2139",
"/TRIM-LINES-2139",
"/PASH-B"
            ]
        }
        return default_objects

    def toggle_side_panel(self):
        """Toggle side panel visibility with animation."""
        if self.panel_visible:
            # Hide panel
            self.animation = QPropertyAnimation(self.side_panel, b"geometry")
            self.animation.setDuration(300)
            self.animation.setStartValue(QRect(700, 0, 300, 800))
            self.animation.setEndValue(QRect(1000, 0, 300, 800))
            self.animation.finished.connect(self.side_panel.hide)
            self.animation.start()
            self.btn_toggle_panel.setText("◄")
            self.btn_toggle_panel.setGeometry(670, 10, 25, 50)
            self.setFixedSize(700, 800)
        else:
            # Show panel
            self.side_panel.show()
            self.animation = QPropertyAnimation(self.side_panel, b"geometry")
            self.animation.setDuration(300)
            self.animation.setStartValue(QRect(1000, 0, 300, 800))
            self.animation.setEndValue(QRect(700, 0, 300, 800))
            self.animation.start()
            self.btn_toggle_panel.setText("►")
            self.btn_toggle_panel.setGeometry(970, 10, 25, 50)
            self.setFixedSize(1000, 800)

        self.panel_visible = not self.panel_visible

    def load_panel_objects(self):
        """Load saved object list for current project from settings."""
        proj_code = self.combo_proj_code.currentText()
        saved_objects = self.settings.value(f'objects_{proj_code}', None)

        if saved_objects:
            self.side_panel.set_items(saved_objects)
        else:
            # Load defaults
            defaults = self.get_default_objects()
            if proj_code in defaults:
                self.side_panel.set_items(defaults[proj_code])
            else:
                self.side_panel.set_items([])

    def save_panel_objects(self):
        """Save current panel objects to settings."""
        proj_code = self.combo_proj_code.currentText()
        objects = self.side_panel.get_items()
        self.settings.setValue(f'objects_{proj_code}', objects)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AppGUI()
    ex.show()
    sys.exit(app.exec())
