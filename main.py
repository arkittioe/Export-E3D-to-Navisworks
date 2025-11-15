# -*- coding: utf-8 -*-

import sys
import os
import json
from pathlib import Path

# Import necessary components from PyQt6
from PyQt6.QtCore import Qt, QSettings, QUrl, QTime
from PyQt6.QtGui import QIcon, QPixmap, QFont, QDesktopServices
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox,
    QFileDialog, QMessageBox, QHBoxLayout, QGroupBox, QStatusBar,
    QDialog, QTextBrowser, QTimeEdit
)


class AppGUI(QMainWindow):
    """
    Main application GUI class for generating AVEVA E3D export files.
    """

    def __init__(self):
        super().__init__()
        self.settings = QSettings('3DDesignMagic', 'E3DExporter')
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

        # Project Code
        self.labels["proj_code"] = QLabel("🏗️ Project Code:")
        self.line_edits["proj_code"] = QLineEdit("PBZ")
        self.line_edits["proj_code"].setToolTip("Enter your project code")
        project_layout.addWidget(self.labels["proj_code"], 0, 0)
        project_layout.addWidget(self.line_edits["proj_code"], 0, 1, 1, 2)

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

        # MDB
        self.labels["mdb"] = QLabel("💾 MDB:")
        self.line_edits["mdb"] = QLineEdit("P2-ALL-PLANT")
        self.line_edits["mdb"].setToolTip("Master Database name")
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
        self.line_edits["object_list"] = QLineEdit("C:/Users/h.izadi/Downloads/SITE.txt")
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
                "proj_code": self.line_edits["proj_code"].text().strip(),
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

            # Basic validation (skip export_time as it can be False)
            for key, value in data.items():
                if key == "export_time":  # Skip export_time validation
                    continue
                if not value and isinstance(value, str):
                    QMessageBox.warning(self, "Input Error",
                                        f"Field '{key.replace('_', ' ').title()}' cannot be empty.")
                    return

            output_folder = Path(data["output_folder"])
            if not output_folder.exists() or not output_folder.is_dir():
                QMessageBox.warning(self, "Path Error", f"Output folder does not exist:\n{output_folder}")
                return

            areas_file_path = Path(data["areas_file"])
            if not areas_file_path.exists() or not areas_file_path.is_file():
                QMessageBox.warning(self, "File Error", f"Object list file not found:\n{areas_file_path}")
                return

            # --- 2. Read Object List from SITE.txt ---
            with open(areas_file_path, 'r', encoding='utf-8') as f:
                object_list = [line.strip() for line in f if line.strip()]

            if not object_list:
                QMessageBox.warning(self, "File Error", "The object list file is empty.")
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
    pause
    exit /b
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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AppGUI()
    ex.show()
    sys.exit(app.exec())
