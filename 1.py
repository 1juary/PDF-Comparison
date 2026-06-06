import re
import sys
import os
import shutil
import tempfile
import unicodedata
from difflib import SequenceMatcher
from json import dump, load
from os import path
from tempfile import TemporaryDirectory
from time import sleep
from typing import Dict, List, Optional, Tuple

import fitz
from PIL import Image, ImageDraw
from PySide6.QtCore import QThread, Signal, Slot, Qt, QPoint
from PySide6.QtGui import QIcon, QPixmap, QImage
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QSpinBox,
    QStyleFactory,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
    QMessageBox,
    QSlider,
)

# ---------------------------------------------------------------------------
# 极简高精度 OCR 引擎 (RapidOCR) 动态装载
# ---------------------------------------------------------------------------
_RAPID_OCR_AVAILABLE = False
_ocr_engine = None

try:
    from rapidocr_onnxruntime import RapidOCR
    _ocr_engine = RapidOCR()
    _RAPID_OCR_AVAILABLE = True
except ImportError:
    pass


# ---------------------------------------------------------------------------
# 配置管理函数
# ---------------------------------------------------------------------------
def resource_path(relative_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return path.join(sys._MEIPASS, relative_path)
    return path.join(path.dirname(path.abspath(__file__)), relative_path)


def _load_default_settings() -> dict:
    return {
        "PAGE_SIZES": {
            "AUTO": [None, None],
            "LETTER": [8.5, 11],
            "ANSI A": [11, 8.5],
            "ANSI B": [17, 11],
            "ANSI C": [22, 17],
            "ANSI D": [34, 22],
        },
        "DPI_LEVELS": [75, 150, 300, 600, 1200, 1800],
        "DPI_LABELS": [
            "Low DPI: Draft Quality [75]",
            "Low DPI: Viewing Quality [150]",
            "Medium DPI: Printable [300]",
            "Standard DPI [600]",
            "High DPI [1200]: Professional Quality",
            "Max DPI [1800]: Large File Size",
        ],
        "INCLUDE_IMAGES": {
            "New Copy": True,
            "Old Copy": True,
            "Markup": True,
            "Difference": False,
            "Overlay": False,
        },
        "DPI": "Standard DPI [600]",
        "DPI_LEVEL": 600,
        "PAGE_SIZE": "AUTO",
        "THRESHOLD": 128,
        "MIN_AREA": 100,
        "EPSILON": 0.0,
        "TEXT_MIN_DIFF_LENGTH": 2,
        "NORMALIZE_TEXT": True,
        "OUTPUT_PATH": None,
        "SCALE_OUTPUT": True,
        "OUTPUT_BW": False,
        "OUTPUT_GS": False,
        "REDUCE_FILESIZE": False,
        "MAIN_PAGE": "New Document",
        "FORCE_OCR": False,
        "VECTOR_BOX_PADDING": 2,  # 矢量标框外延拓展像素
        "OCR_BOX_PADDING": 2,     # OCR 标框外延拓展像素
        "OCR_MERGE_DIST_H": 15,   # OCR 左右拼接物理临界阈值
    }


def _normalize_settings(settings: dict) -> dict:
    defaults = _load_default_settings()

    for key, value in defaults.items():
        if key not in settings:
            settings[key] = value
            continue

        if isinstance(value, dict) and isinstance(settings[key], dict):
            for child_key, child_default in value.items():
                settings[key].setdefault(child_key, child_default)

    if isinstance(settings.get("PAGE_SIZE"), list):
        page_size_list = settings["PAGE_SIZE"]
        matched = "AUTO"
        for name, size in settings["PAGE_SIZES"].items():
            if list(size) == list(page_size_list):
                matched = name
                break
        settings["PAGE_SIZE"] = matched

    if settings.get("PAGE_SIZE") not in settings.get("PAGE_SIZES", {}):
        settings["PAGE_SIZE"] = "AUTO"

    return settings


def load_settings() -> dict:
    settings_path = "settings.json"
    settings = None
    if path.exists(settings_path):
        with open(settings_path, "r", encoding="utf-8") as file:
            settings = load(file)

    if not settings:
        settings = _load_default_settings()

    settings = _normalize_settings(settings)
    save_settings(settings)
    return settings


def save_settings(settings: dict) -> None:
    settings_path = "settings.json"
    with open(settings_path, "w", encoding="utf-8") as file:
        dump(settings, file, indent=4)


# ---------------------------------------------------------------------------
# 设置子页面组件
# ---------------------------------------------------------------------------
class AdvancedSettings(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = load_settings()

        self.min_diff_label = QLabel("Minimum Diff Token Length [Default: 2]:")
        self.min_diff_desc = QLabel(
            "Short differences are usually noise in PDFs. Increase this value to suppress tiny token-level changes."
        )
        self.min_diff_desc.setWordWrap(True)
        self.min_diff_desc.setStyleSheet("color: #6B7280; font: 12px 'Segoe UI', Arial, sans-serif;")
        self.min_diff_spinbox = QSpinBox(self)
        self.min_diff_spinbox.setMinimum(1)
        self.min_diff_spinbox.setMaximum(20)
        self.min_diff_spinbox.setValue(self.settings.get("TEXT_MIN_DIFF_LENGTH", 2))
        self.min_diff_spinbox.valueChanged.connect(self.update_min_diff)

        self.normalize_checkbox = QCheckBox("Normalize Text (lowercase, trim punctuation)")
        self.normalize_checkbox.setChecked(self.settings.get("NORMALIZE_TEXT", True))
        self.normalize_checkbox.stateChanged.connect(self.update_normalize)

        self.force_ocr_checkbox = QCheckBox("Force OCR on all pages (强制所有页面进行OCR)")
        self.force_ocr_checkbox.setChecked(self.settings.get("FORCE_OCR", False))
        self.force_ocr_checkbox.stateChanged.connect(self.update_force_ocr)

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.addWidget(self.min_diff_label)
        layout.addWidget(self.min_diff_desc)
        layout.addWidget(self.min_diff_spinbox)
        layout.addWidget(self.normalize_checkbox)
        layout.addWidget(self.force_ocr_checkbox)
        self.setLayout(layout)

        self.setStyleSheet("""
            QLabel {
                color: #1A1A2E;
                font: 14px "Segoe UI", Arial, sans-serif;
            }
            QSpinBox {
                color: #1A1A2E;
                font: 14px "Segoe UI", Arial, sans-serif;
                background-color: white;
                border: 1px solid #E0E6ED;
                border-radius: 6px;
                padding: 4px;
            }
            QSpinBox:focus {
                border: 1px solid #2196F3;
            }
            QCheckBox {
                color: #1A1A2E;
                font: 14px "Segoe UI", Arial, sans-serif;
                margin-top: 4px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 2px solid #E0E6ED;
            }
            QCheckBox::indicator:checked {
                background-color: #2196F3;
                border: 2px solid #2196F3;
            }
        """)

    def update_min_diff(self, value):
        self.settings["TEXT_MIN_DIFF_LENGTH"] = int(value)
        save_settings(self.settings)

    def update_normalize(self, state):
        self.settings["NORMALIZE_TEXT"] = state == 2
        save_settings(self.settings)

    def update_force_ocr(self, state):
        self.settings["FORCE_OCR"] = state == 2
        save_settings(self.settings)


class DPISettings(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = None
        
        # 递归寻找父层MainWindow
        curr = parent
        while curr:
            if isinstance(curr, QMainWindow):
                self.parent_window = curr
                break
            curr = curr.parent()
            
        self.settings = load_settings()

        self.low_draft_label = QLabel("Low DPI - Draft Quality:")
        self.low_draft_spinbox = QSpinBox(self)
        self.low_draft_spinbox.setMinimum(1)
        self.low_draft_spinbox.setMaximum(99)
        self.low_draft_spinbox.setValue(self.settings["DPI_LEVELS"][0])
        self.low_draft_spinbox.valueChanged.connect(self.update_dpi_levels)

        self.low_viewing_label = QLabel("Low DPI - Viewing Only:")
        self.low_viewing_spinbox = QSpinBox(self)
        self.low_viewing_spinbox.setMinimum(100)
        self.low_viewing_spinbox.setMaximum(199)
        self.low_viewing_spinbox.setValue(self.settings["DPI_LEVELS"][1])
        self.low_viewing_spinbox.valueChanged.connect(self.update_dpi_levels)

        self.medium_label = QLabel("Medium DPI - Printable:")
        self.medium_spinbox = QSpinBox(self)
        self.medium_spinbox.setMinimum(200)
        self.medium_spinbox.setMaximum(599)
        self.medium_spinbox.setValue(self.settings["DPI_LEVELS"][2])
        self.medium_spinbox.valueChanged.connect(self.update_dpi_levels)

        self.standard_label = QLabel("Standard DPI:")
        self.standard_spinbox = QSpinBox(self)
        self.standard_spinbox.setMinimum(600)
        self.standard_spinbox.setMaximum(999)
        self.standard_spinbox.setValue(self.settings["DPI_LEVELS"][3])
        self.standard_spinbox.valueChanged.connect(self.update_dpi_levels)

        self.high_label = QLabel("High DPI - Professional Quality:")
        self.high_spinbox = QSpinBox(self)
        self.high_spinbox.setMinimum(1000)
        self.high_spinbox.setMaximum(1999)
        self.high_spinbox.setValue(self.settings["DPI_LEVELS"][4])
        self.high_spinbox.valueChanged.connect(self.update_dpi_levels)

        self.max_label = QLabel("Max DPI - Large File Size:")
        self.max_spinbox = QSpinBox(self)
        self.max_spinbox.setMinimum(1000)
        self.max_spinbox.setMaximum(6000)
        self.max_spinbox.setValue(self.settings["DPI_LEVELS"][5])
        self.max_spinbox.valueChanged.connect(self.update_dpi_levels)

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.addWidget(self.low_draft_label)
        layout.addWidget(self.low_draft_spinbox)
        layout.addWidget(self.low_viewing_label)
        layout.addWidget(self.low_viewing_spinbox)
        layout.addWidget(self.medium_label)
        layout.addWidget(self.medium_spinbox)
        layout.addWidget(self.standard_label)
        layout.addWidget(self.standard_spinbox)
        layout.addWidget(self.high_label)
        layout.addWidget(self.high_spinbox)
        layout.addWidget(self.max_label)
        layout.addWidget(self.max_spinbox)
        self.setLayout(layout)

        self.setStyleSheet("""
            QLabel {
                color: #1A1A2E;
                font: 14px "Segoe UI", Arial, sans-serif;
            }
            QSpinBox {
                color: #1A1A2E;
                font: 14px "Segoe UI", Arial, sans-serif;
                background-color: white;
                border: 1px solid #E0E6ED;
                border-radius: 6px;
                padding: 4px;
            }
            QSpinBox:focus {
                border: 1px solid #2196F3;
            }
        """)

    def update_dpi_levels(self, new_dpi):
        if new_dpi < 100:
            self.settings["DPI_LEVELS"][0] = new_dpi
            self.settings["DPI_LABELS"][0] = f"Low DPI: Draft Quality [{new_dpi}]"
        elif new_dpi < 200:
            self.settings["DPI_LEVELS"][1] = new_dpi
            self.settings["DPI_LABELS"][1] = f"Low DPI: Viewing Quality [{new_dpi}]"
        elif new_dpi < 600:
            self.settings["DPI_LEVELS"][2] = new_dpi
            self.settings["DPI_LABELS"][2] = f"Medium DPI: Printable [{new_dpi}]"
        elif new_dpi < 1000:
            self.settings["DPI_LEVELS"][3] = new_dpi
            self.settings["DPI_LABELS"][3] = f"Standard DPI [{new_dpi}]"
        elif new_dpi < 2000:
            self.settings["DPI_LEVELS"][4] = new_dpi
            self.settings["DPI_LABELS"][4] = f"High DPI: Professional Quality [{new_dpi}]"
        else:
            self.settings["DPI_LEVELS"][5] = new_dpi
            self.settings["DPI_LABELS"][5] = f"Max DPI: High Memory [{new_dpi}]"

        save_settings(self.settings)
        if self.parent_window and hasattr(self.parent_window, "dpi_combo"):
            self.parent_window.dpi_combo.clear()
            self.parent_window.dpi_combo.addItems(self.settings["DPI_LABELS"])
            self.parent_window.dpi_combo.setCurrentText(self.settings["DPI_LABELS"][3])


class OutputSettings(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = load_settings()
        self.output_path_label = QLabel("Output Path:")
        self.output_path_combobox = QComboBox(self)
        self.output_path_combobox.addItems(["Source Path", "Default Path", "Specified Path"])
        if self.settings["OUTPUT_PATH"] == "\\":
            self.output_path_combobox.setCurrentText("Default Path")
        elif self.settings["OUTPUT_PATH"] is None:
            self.output_path_combobox.setCurrentText("Source Path")
        else:
            self.output_path_combobox.setCurrentText("Specified Path")
        self.output_path_combobox.currentTextChanged.connect(self.set_output_path)

        self.specified_label = QLabel("Specified Path:")
        self.specified_entry = QLineEdit(self)
        self.specified_entry.setText(
            self.settings["OUTPUT_PATH"] if self.output_path_combobox.currentText() == "Specified Path" else ""
        )
        self.specified_entry.textChanged.connect(self.set_output_path)

        self.checkbox_image1 = QCheckBox("New Copy")
        self.checkbox_image1.setChecked(self.settings["INCLUDE_IMAGES"]["New Copy"])
        self.checkbox_image2 = QCheckBox("Old Copy")
        self.checkbox_image2.setChecked(self.settings["INCLUDE_IMAGES"]["Old Copy"])
        self.checkbox_image3 = QCheckBox("Markup")
        self.checkbox_image3.setChecked(self.settings["INCLUDE_IMAGES"]["Markup"])
        self.checkbox_image4 = QCheckBox("Difference")
        self.checkbox_image4.setChecked(self.settings["INCLUDE_IMAGES"]["Difference"])
        self.checkbox_image5 = QCheckBox("Overlay")
        self.checkbox_image5.setChecked(self.settings["INCLUDE_IMAGES"]["Overlay"])

        self.checkbox_image1.stateChanged.connect(self.set_output_images)
        self.checkbox_image2.stateChanged.connect(self.set_output_images)
        self.checkbox_image3.stateChanged.connect(self.set_output_images)
        self.checkbox_image4.stateChanged.connect(self.set_output_images)
        self.checkbox_image5.stateChanged.connect(self.set_output_images)

        self.scaling_checkbox = QCheckBox("Scale Pages")
        self.scaling_checkbox.setChecked(self.settings["SCALE_OUTPUT"])
        self.scaling_checkbox.stateChanged.connect(self.set_scaling)

        self.bw_checkbox = QCheckBox("Black/White")
        self.bw_checkbox.setChecked(self.settings["OUTPUT_BW"])
        self.bw_checkbox.stateChanged.connect(self.set_bw)

        self.gs_checkbox = QCheckBox("Grayscale")
        self.gs_checkbox.setChecked(self.settings["OUTPUT_GS"])
        self.gs_checkbox.stateChanged.connect(self.set_gs)

        self.reduce_checkbox = QCheckBox("Reduce Size")
        self.reduce_checkbox.setChecked(self.settings["REDUCE_FILESIZE"])
        self.reduce_checkbox.stateChanged.connect(self.set_reduced_filesize)

        self.main_page_label = QLabel("Main Page:")
        self.main_page_combobox = QComboBox(self)
        self.main_page_combobox.addItems(["New Document", "Old Document"])
        self.main_page_combobox.setCurrentText(self.settings["MAIN_PAGE"])
        self.main_page_combobox.currentTextChanged.connect(self.set_main_page)

        output_path_group = QGroupBox("Output Settings")
        include_images_group = QGroupBox("Files to include:")
        general_group = QGroupBox("General")
        checkboxes_group = QGroupBox()
        other_group = QGroupBox()

        output_path_layout = QFormLayout()
        output_path_layout.addRow(self.output_path_label, self.output_path_combobox)
        output_path_layout.addRow(self.specified_label, self.specified_entry)
        output_path_group.setLayout(output_path_layout)

        include_images_layout = QHBoxLayout()
        include_images_layout.addWidget(self.checkbox_image1)
        include_images_layout.addWidget(self.checkbox_image2)
        include_images_layout.addWidget(self.checkbox_image3)
        include_images_layout.addWidget(self.checkbox_image4)
        include_images_layout.addWidget(self.checkbox_image5)
        include_images_group.setLayout(include_images_layout)

        general_layout = QHBoxLayout()
        checkboxes = QVBoxLayout()
        other = QVBoxLayout()
        other.setAlignment(Qt.AlignmentFlag.AlignTop)
        checkboxes.addWidget(self.scaling_checkbox)
        checkboxes.addWidget(self.bw_checkbox)
        checkboxes.addWidget(self.gs_checkbox)
        checkboxes.addWidget(self.reduce_checkbox)
        other.addWidget(self.main_page_label)
        other.addWidget(self.main_page_combobox)
        checkboxes_group.setLayout(checkboxes)
        other_group.setLayout(other)
        general_layout.addWidget(checkboxes_group)
        general_layout.addWidget(other_group)
        general_group.setLayout(general_layout)

        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        main_layout.addWidget(output_path_group)
        main_layout.addWidget(include_images_group)
        main_layout.addWidget(general_group)
        self.setLayout(main_layout)

        self.setStyleSheet("""
            QLabel {
                color: #1A1A2E;
                font: 14px "Segoe UI", Arial, sans-serif;
            }
            QLineEdit {
                color: #1A1A2E;
                background-color: white;
                border: 1px solid #E0E6ED;
                border-radius: 6px;
                padding: 6px;
                font: 13px "Segoe UI", Arial, sans-serif;
            }
            QLineEdit:focus {
                border: 1px solid #2196F3;
            }
            QComboBox {
                height: 32px;
                border-radius: 8px;
                background-color: white;
                border: 1px solid #E0E6ED;
                color: #1A1A2E;
                padding: 4px 8px;
                font: 13px "Segoe UI", Arial, sans-serif;
            }
            QComboBox:hover {
                border: 1px solid #2196F3;
            }
            QComboBox QAbstractItemView {
                padding: 6px;
                background-color: white;
                selection-background-color: #E3F2FD;
                selection-color: #1565C0;
                color: #1A1A2E;
                font: 13px "Segoe UI", Arial, sans-serif;
            }
            QComboBox::drop-down {
                border: none;
            }
            QCheckBox {
                color: #1A1A2E;
                font: 14px "Segoe UI", Arial, sans-serif;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 2px solid #E0E6ED;
            }
            QCheckBox::indicator:checked {
                background-color: #2196F3;
                border: 2px solid #2196F3;
            }
            QGroupBox {
                color: #1A1A2E;
                font: bold 14px "Segoe UI", Arial, sans-serif;
                border: 1px solid #E0E6ED;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px 0 4px;
                color: #2196F3;
            }
        """)

    def set_output_path(self, option):
        if option == "Source Path":
            self.settings["OUTPUT_PATH"] = None
        elif option == "Default Path":
            self.settings["OUTPUT_PATH"] = "\\"
        else:
            raw = self.specified_entry.text().strip()
            if raw:
                normalized = raw.replace("/", "\\")
                if not normalized.endswith("\\"):
                    normalized += "\\"
                self.settings["OUTPUT_PATH"] = normalized
            else:
                self.settings["OUTPUT_PATH"] = None

        save_settings(self.settings)

    def set_output_images(self, state):
        checkbox = self.sender()
        self.settings["INCLUDE_IMAGES"][checkbox.text()] = state == 2
        save_settings(self.settings)

    def set_scaling(self, state):
        self.settings["SCALE_OUTPUT"] = state == 2
        save_settings(self.settings)

    def set_bw(self, state):
        self.settings["OUTPUT_BW"] = state == 2
        save_settings(self.settings)

    def set_gs(self, state):
        self.settings["OUTPUT_GS"] = state == 2
        save_settings(self.settings)

    def set_reduced_filesize(self, state):
        self.settings["REDUCE_FILESIZE"] = state == 2
        save_settings(self.settings)

    def set_main_page(self, page):
        self.settings["MAIN_PAGE"] = page
        save_settings(self.settings)


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setWindowModality(Qt.ApplicationModal)
        self.setFixedSize(520, 550)

        self.tab_widget = QTabWidget(self)
        self.tab_widget.addTab(OutputSettings(self), "Output")
        self.tab_widget.addTab(DPISettings(self), "DPI")
        self.tab_widget.addTab(AdvancedSettings(self), "Advanced")

        layout = QVBoxLayout(self)
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

        self.setStyleSheet("""
            QDialog {
                background-color: #FAFBFC;
                color: #1A1A2E;
            }
            QTabWidget::pane {
                border: 1px solid #E0E6ED;
                background: white;
                border-radius: 8px;
            }
            QTabBar::tab {
                background: #F0F4F8;
                color: #6B7280;
                padding: 10px 24px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 2px;
                font: 13px "Segoe UI", Arial, sans-serif;
            }
            QTabBar::tab:selected {
                background: #2196F3;
                color: white;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background: #E3F2FD;
                color: #1565C0;
            }
        """)


# ---------------------------------------------------------------------------
# 自定义标题栏与基础按钮
# ---------------------------------------------------------------------------
class CustomTitleBar(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setFixedHeight(40)
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(10, 0, 10, 0)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        spacer_item = QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.title_label = QLabel("PDF Comparison Tool")
        self.title_label.setStyleSheet("color: white; font: bold 14px 'Segoe UI', Arial, sans-serif;")
        self.settings_button = QPushButton("Settings", self)
        self.settings_button.setObjectName("SettingsButton")
        self.settings_button.setFixedSize(70, 28)
        self.settings_button.clicked.connect(self.open_settings)

        self.minimize_button = QPushButton("−", self)
        self.minimize_button.setObjectName("MinimizeButton")
        self.minimize_button.setFixedSize(28, 28)
        self.minimize_button.clicked.connect(self.parent.showMinimized)

        self.close_button = QPushButton("✕", self)
        self.close_button.setObjectName("CloseButton")
        self.close_button.setFixedSize(28, 28)
        self.close_button.clicked.connect(self.parent.close)

        self.layout.addWidget(self.settings_button)
        self.layout.addItem(spacer_item)
        self.title_label.setText("PDF Comparison Tool")
        self.layout.addWidget(self.title_label)
        self.layout.addItem(spacer_item)
        self.layout.addWidget(self.minimize_button)
        self.layout.addWidget(self.close_button)

        self.draggable = True
        self.dragging_threshold = 5
        self.drag_start_position = None

    def mousePressEvent(self, event):
        if self.draggable and event.button() == Qt.LeftButton:
            self.drag_start_position = event.globalPosition().toPoint() - self.parent.pos()
        event.accept()

    def mouseMoveEvent(self, event):
        if self.draggable and self.drag_start_position is not None and event.buttons() == Qt.LeftButton:
            if (event.globalPosition().toPoint() - self.drag_start_position).manhattanLength() > self.dragging_threshold:
                self.parent.move(event.globalPosition().toPoint() - self.drag_start_position)
                self.drag_start_position = event.globalPosition().toPoint() - self.parent.pos()
        event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = None
        event.accept()

    def open_settings(self):
        settings_dialog = SettingsDialog(self.parent)
        settings_dialog.exec()


class DragDropLabel(QPushButton):
    def __init__(self, role, parent=None):
        super().__init__(parent)
        self._parent = parent
        self.role = role  # "old" or "new"
        self.file_path = None
        self.setAcceptDrops(True)
        size_policy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.clicked.connect(self.browse_file)
        self.setSizePolicy(size_policy)
        self.setMinimumHeight(100)
        self._update_style()

    def _update_style(self):
        if self.role == "old":
            color = "#FF6B6B"
            bg_color = "#FFF5F5"
            bg_hover = "#FFE0E0"
            label = "Drop OLD Version PDF here\nor click to browse"
        else:
            color = "#2196F3"
            bg_color = "#F0F8FF"
            bg_hover = "#E3F2FD"
            label = "Drop NEW Version PDF here\nor click to browse"

        if self.file_path:
            label = f"{'OLD' if self.role == 'old' else 'NEW'}: {path.basename(self.file_path)}"

        self.setText(label)
        self.setStyleSheet(f"""
        QPushButton {{
            color: {color};
            background-color: {bg_color};
            border-radius: 12px;
            border: 2px dashed {color};
            font: bold 13px "Segoe UI", Arial, sans-serif;
            padding: 8px;
        }}
        QPushButton:hover {{
            background-color: {bg_hover};
            border: 2px solid {color};
        }}
        """)

    def browse_file(self):
        file = QFileDialog.getOpenFileName(
            self,
            f"Open {'Old Version' if self.role == 'old' else 'New Version'} PDF",
            "",
            "PDF Files (*.pdf)"
        )[0]
        if file:
            self.set_file(file)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(".pdf"):
                self.set_file(file_path)

    def set_file(self, file_path):
        self.file_path = file_path
        if self._parent.files is None:
            self._parent.files = [None, None]
        if self.role == "new":
            self._parent.files[0] = file_path
        else:
            self._parent.files[1] = file_path
        self._update_style()
        if hasattr(self._parent, "update_ui_state"):
            self._parent.update_ui_state()


# ---------------------------------------------------------------------------
# 进度查看子窗口
# ---------------------------------------------------------------------------
class ProgressWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyPDFCompare Progress")
        self.resize(600, 500)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()

        self.progress_bar = QProgressBar()
        self.log_area = QTextBrowser()
        self.log_area.setReadOnly(True)

        self.layout.addWidget(self.progress_bar)
        self.layout.addWidget(self.log_area)
        self.central_widget.setLayout(self.layout)

        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #FAFBFC;
            }
            QTextBrowser {
                background-color: white;
                color: #1A1A2E;
                border: 1px solid #E0E6ED;
                border-radius: 8px;
                font: 12px "Segoe UI", Arial, sans-serif;
                padding: 8px;
            }
            QProgressBar {
                border: 1px solid #E0E6ED;
                border-radius: 8px;
                text-align: center;
                color: #1A1A2E;
                background-color: #F0F4F8;
                height: 20px;
                font: 11px "Segoe UI", Arial, sans-serif;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
                border-radius: 7px;
            }
        """
        )

    @Slot(int)
    def update_progress(self, progress):
        self.progress_bar.setValue(progress)

    @Slot(str)
    def update_log(self, message):
        self.log_area.append(message)

    @Slot(int)
    def operation_complete(self, delay_seconds):
        sleep(delay_seconds)
        self.close()


# ---------------------------------------------------------------------------
# 语义理解高能比对子线程 (核心比对逻辑)
# ---------------------------------------------------------------------------
class CompareThread(QThread):
    progressUpdated = Signal(int)
    compareComplete = Signal(int)
    logMessage = Signal(str)

    def __init__(self, files: List[str], progress_window: ProgressWindow, parent=None):
        super().__init__(parent)
        compare_settings = load_settings()

        self.DPI_LEVEL = compare_settings.get("DPI_LEVEL", 600)
        self.PAGE_SIZE_NAME = compare_settings.get("PAGE_SIZE", "AUTO")
        self.PAGE_SIZES = compare_settings.get("PAGE_SIZES", {})
        self.PAGE_SIZE = tuple(self.PAGE_SIZES.get(self.PAGE_SIZE_NAME, [None, None]))
        self.INCLUDE_IMAGES = compare_settings.get("INCLUDE_IMAGES", {})
        self.MAIN_PAGE = compare_settings.get("MAIN_PAGE", "New Document")
        self.OUTPUT_PATH = compare_settings.get("OUTPUT_PATH")
        self.SCALE_OUTPUT = compare_settings.get("SCALE_OUTPUT", True)
        self.OUTPUT_BW = compare_settings.get("OUTPUT_BW", False)
        self.OUTPUT_GS = compare_settings.get("OUTPUT_GS", False)
        self.REDUCE_FILESIZE = compare_settings.get("REDUCE_FILESIZE", False)
        self.TEXT_MIN_DIFF_LENGTH = int(compare_settings.get("TEXT_MIN_DIFF_LENGTH", 2))
        self.NORMALIZE_TEXT = bool(compare_settings.get("NORMALIZE_TEXT", True))
        self.FORCE_OCR = bool(compare_settings.get("FORCE_OCR", False))

        # 用户通过滑动条调节的框大小与合并范围数值
        self.VECTOR_BOX_PADDING = int(compare_settings.get("VECTOR_BOX_PADDING", 2))
        self.OCR_BOX_PADDING = int(compare_settings.get("OCR_BOX_PADDING", 2))
        self.OCR_MERGE_DIST_H = int(compare_settings.get("OCR_MERGE_DIST_H", 15))

        self.files = files
        self.progress_window = progress_window
        self.statistics = {
            "NUM_PAGES": 0,
            "MAIN_PAGE": None,
            "TOTAL_DIFFERENCES": 0,
            "PAGES_WITH_DIFFERENCES": [],
            "ADDED_COUNT": 0,
            "DELETED_COUNT": 0,
        }

        self.progressUpdated.connect(self.progress_window.update_progress)
        self.logMessage.connect(self.progress_window.update_log)
        self.compareComplete.connect(self.progress_window.operation_complete)

    def run(self):
        try:
            self.handle_files(self.files)
        except fitz.fitz.FileDataError as error:
            self.logMessage.emit(f"Error opening file: {error}")
        except Exception as error:
            self.logMessage.emit(f"Unhandled comparison error: {error}")

    def _normalize_text(self, text: str) -> str:
        """
        高保真度语义归一化转换
        """
        if not text:
            return ""
        
        # 1. Unicode规范化：分解复合字、半角/全角自动映射
        text = unicodedata.normalize('NFKC', text)
        
        # 2. 小数前导零还原（如：将 '±.5' 转换为 '±0.5'，'.5' 转换为 '0.5'）
        text = re.sub(r'(?<!\d)\.(\d+)', r'0.\1', text)
        text = text.strip()
        
        if self.NORMALIZE_TEXT:
            text = text.lower()
            # 3. 剔除不可见排版干扰符 (软连字符、零宽符号等)
            text = re.sub(r'[\xad\u200b\u200c\u200d\ufeff]+', '', text)
            # 4. 彻底压缩所有行内空白与断行
            text = re.sub(r"[\s\t\r\n]+", "", text)
            # 5. 过滤标点、图纸干扰标识、特殊工业符号（★☆等）
            text = re.sub(r"[\.,;:()\[\]{}<>\-_=+`~\"'“”‘’★☆●■▲▼◆○✓✔✗✘✕✖+-]+", "", text)
            
        return text

    def _merge_close_tokens(self, tokens: List[Dict]) -> List[Dict]:
        """
        高能空间物理邻近片段判定合并 (全面防平行行粘连版)
        """
        if not tokens:
            return []
        
        n = len(tokens)
        adj = {i: [] for i in range(n)}
        dist_h = float(self.OCR_MERGE_DIST_H)  # 对应界面的 OCR 拼接阈值
        dist_v = 6.0
        
        for i in range(n):
            r1 = tokens[i]['rect']
            for j in range(i + 1, n):
                r2 = tokens[j]['rect']
                
                # 判定垂直方向是否重叠/相近
                v_overlap = max(r1.y0 - dist_v, r2.y0) < min(r1.y1 + dist_v, r2.y1)
                
                # 判定水平贴近边缘
                h_close = False
                if v_overlap:
                    # 首先确认在水平范围内
                    if r1.x1 + dist_h >= r2.x0 and r1.x0 <= r2.x1 + dist_h:
                        # 倾斜行平行安全检查：如果它们在水平上重合过大，说明它们是并排平行的不同标注，禁止合并！
                        h_overlap = min(r1.x1, r2.x1) - max(r1.x0, r2.x0)
                        if h_overlap <= 5.0:
                            h_close = True
                        
                if v_overlap and h_close:
                    adj[i].append(j)
                    adj[j].append(i)
                    
        visited = [False] * n
        merged_tokens = []
        
        for i in range(n):
            if not visited[i]:
                component = []
                queue = [i]
                visited[i] = True
                while queue:
                    curr = queue.pop(0)
                    component.append(tokens[curr])
                    for neighbor in adj[curr]:
                        if not visited[neighbor]:
                            visited[neighbor] = True
                            queue.append(neighbor)
                            
                # 按 X 坐标从左往右排序合并
                component.sort(key=lambda t: t['rect'].x0)
                
                merged_text = " ".join(t['text'].strip() for t in component)
                union_rect = component[0]['rect']
                for t in component[1:]:
                    union_rect = union_rect | t['rect']
                    
                merged_tokens.append({
                    "text": merged_text,
                    "norm": self._normalize_text(merged_text),
                    "page": component[0]['page'],
                    "rect": union_rect
                })
                
        return merged_tokens

    def _extract_tokens(self, doc: fitz.Document) -> List[Dict]:
        """
        物理布局分行 (Layout-based Grouping) 与 延迟长度过滤高精度提取
        """
        tokens = []
        
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            
            # 1. 优先提取物理排版矢量词信息
            words = page.get_text("words")
            
            # 判定有效矢量词数是否过少（如果是扫描件或者非完整矢量图纸）或者是否开启了强制OCR
            real_word_count = sum(1 for w in words if (w[4] or "").strip())
            is_scanned = (real_word_count < 12) or self.FORCE_OCR
            
            # 2. 如果页面无法提取到充足字符，或用户开启了强制OCR，启用 OCR 模块修复
            if is_scanned:
                self.logMessage.emit(f"Page {page_num + 1}: Scanned page/low-text page detected (Force OCR={self.FORCE_OCR}, Word count={real_word_count}). Attempting OCR...")
                ocr_words = []
                if _RAPID_OCR_AVAILABLE:
                    try:
                        dpi = 150
                        pix = page.get_pixmap(dpi=dpi)
                        # 将 Pixmap 转换为 PIL Image 以获得最高兼容性
                        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                        
                        result, _ = _ocr_engine(img)
                        if result:
                            scale = dpi / 72.0
                            for idx, item in enumerate(result):
                                box, text, conf = item
                                xs = [pt[0] for pt in box]
                                ys = [pt[1] for pt in box]
                                px0, py0, px2, py2 = min(xs), min(ys), max(xs), max(ys)
                                x0, y0, x2, y2 = px0 / scale, py0 / scale, px2 / scale, py2 / scale
                                # 每个 OCR 行单独归入对应的唯一 block_no
                                ocr_words.append((x0, y0, x2, y2, text, idx, 0, 0))
                                
                        if len(ocr_words) > 0:
                            words = ocr_words
                            self.logMessage.emit(f"Page {page_num + 1}: RapidOCR applied successfully ({len(words)} lines found).")
                        else:
                            self.logMessage.emit(f"Page {page_num + 1}: RapidOCR ran but returned 0 results.")
                    except Exception as e:
                        self.logMessage.emit(f"Page {page_num + 1}: RapidOCR execution failed: {e}")
                else:
                    self.logMessage.emit(f"Page {page_num + 1}: RapidOCR is NOT available. Falling back to Tesseract...")
                    try:
                        tp = page.get_textpage_ocr(flags=0, language="chi_sim+eng", dpi=150, full=True)
                        tess_words = page.get_text("words", textpage=tp)
                        if len(tess_words) > 0:
                            words = tess_words
                            self.logMessage.emit(f"Page {page_num + 1}: Tesseract fallback applied successfully.")
                        else:
                            self.logMessage.emit(f"Page {page_num + 1}: Tesseract fallback returned 0 results.")
                    except Exception as ocr_err:
                        self.logMessage.emit(
                            f"Page {page_num + 1}: [WARNING] No OCR engine available! Text in this scanned/image page cannot be compared."
                        )

            # 3. 物理行分组：先完整保留字符块（不进行极短字符过滤），防止碎化CAD文字被提早过滤
            if words:
                groups = {}
                for word in words:
                    raw = (word[4] or "").strip()
                    if not raw:
                        continue
                    
                    # 仅根据 block_no 和 line_no 做初始归类
                    block_no = word[5]
                    line_no = word[6]
                    key = (block_no, line_no)
                    groups.setdefault(key, []).append(word)
                
                initial_tokens = []
                for (b_no, l_no), g_words in groups.items():
                    # 依据行内词序号 word_no 进行阅读顺序还原
                    g_words.sort(key=lambda w: w[7])
                    
                    merged_text = " ".join(w[4].strip() for w in g_words)
                    
                    x0 = min(w[0] for w in g_words)
                    y0 = min(w[1] for w in g_words)
                    x1 = max(w[2] for w in g_words)
                    y1 = max(w[3] for w in g_words)
                    rect = fitz.Rect(x0, y0, x1, y1)
                    
                    initial_tokens.append({
                        "text": merged_text,
                        "norm": self._normalize_text(merged_text),
                        "page": page_num,
                        "rect": rect
                    })
                
                # 4. 空间合并：对所有图纸文字统一运行安全合并（把打碎的 CAD 单字拼接起来）
                merged_page_tokens = self._merge_close_tokens(initial_tokens)
                
                # 5. 延迟长度过滤：在单字彻底拼装为整行特征后，才根据设置进行长度滤除！
                final_page_tokens = []
                for token in merged_page_tokens:
                    if len(token["norm"]) < self.TEXT_MIN_DIFF_LENGTH:
                        continue
                    final_page_tokens.append(token)
                    
                tokens.extend(final_page_tokens)
                
        return tokens

    @staticmethod
    def _tokens_to_text(tokens: List[Dict], max_tokens: int = 18) -> str:
        if not tokens:
            return "无"
        text = " ".join(token["text"] for token in tokens[:max_tokens]).strip()
        if len(tokens) > max_tokens:
            text += " ..."
        return text

    @staticmethod
    def _group_rects_by_page(tokens: List[Dict]) -> Dict[int, List[fitz.Rect]]:
        rect_map: Dict[int, List[fitz.Rect]] = {}
        for token in tokens:
            rect_map.setdefault(token["page"], []).append(token["rect"])
        return rect_map

    def _build_diff_entries(self, old_tokens: List[Dict], new_tokens: List[Dict]) -> List[Dict]:
        old_norm = [token["norm"] for token in old_tokens]
        new_norm = [token["norm"] for token in new_tokens]
        matcher = SequenceMatcher(None, old_norm, new_norm, autojunk=False)

        entries = []
        for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
            if opcode == "equal":
                continue

            old_slice = old_tokens[i1:i2]
            new_slice = new_tokens[j1:j2]
            old_desc = self._tokens_to_text(old_slice)
            new_desc = self._tokens_to_text(new_slice)

            if old_desc == "无" and new_desc == "无":
                continue

            entry_type = "replace" if opcode == "replace" else ("delete" if opcode == "delete" else "add")
            old_rects = self._group_rects_by_page(old_slice)
            new_rects = self._group_rects_by_page(new_slice)

            entry = {
                "type": entry_type,
                "old_desc": old_desc,
                "old_page": (old_slice[0]["page"] + 1) if old_slice else "无",
                "new_desc": new_desc,
                "new_page": (new_slice[0]["page"] + 1) if new_slice else "无",
                "old_rects": old_rects,
                "new_rects": new_rects,
            }
            entries.append(entry)

        return entries

    @staticmethod
    def _render_page(doc: fitz.Document, page_index: int, dpi: int) -> Tuple[Image.Image, fitz.Rect]:
        if page_index < doc.page_count:
            page = doc.load_page(page_index)
            pix = page.get_pixmap(dpi=dpi)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            return image, page.rect

        first_page = doc.load_page(0)
        pix = first_page.get_pixmap(dpi=dpi)
        return Image.new("RGB", (pix.width, pix.height), (255, 255, 255)), first_page.rect

    @staticmethod
    def _draw_rectangles(image: Image.Image, page_rect: fitz.Rect, rects: List[fitz.Rect], color: Tuple[int, int, int], padding: int = 2):
        if not rects:
            return image

        x_scale = image.width / max(page_rect.width, 1)
        y_scale = image.height / max(page_rect.height, 1)
        stroke = max(1, int(min(image.width, image.height) / 800))

        overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
        overlay_draw = ImageDraw.Draw(overlay)
        fill_color = (*color, 50)
        outline_color = (*color, 160)

        for rect in rects:
            x0 = max(0, int(rect.x0 * x_scale) - padding)
            y0 = max(0, int(rect.y0 * y_scale) - padding)
            x1 = min(image.width - 1, int(rect.x1 * x_scale) + padding)
            y1 = min(image.height - 1, int(rect.y1 * y_scale) + padding)
            overlay_draw.rectangle((x0, y0, x1, y1), fill=fill_color, outline=outline_color, width=stroke)

        image_rgba = image.convert("RGBA")
        composited = Image.alpha_composite(image_rgba, overlay)
        return composited.convert("RGB")

    def _resize_if_needed(self, image: Image.Image) -> Image.Image:
        if not self.SCALE_OUTPUT:
            return image

        if self.PAGE_SIZE[0] is None or self.PAGE_SIZE[1] is None:
            return image

        return image.resize((int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL)))

    @staticmethod
    def _combine_side_by_side(left: Image.Image, right: Image.Image) -> Image.Image:
        height = max(left.height, right.height)
        width = left.width + right.width
        merged = Image.new("RGB", (width, height), (255, 255, 255))
        merged.paste(left, (0, 0))
        merged.paste(right, (left.width, 0))
        return merged

    @staticmethod
    def _overlay_blend(base: Image.Image, other: Image.Image) -> Image.Image:
        other = other.resize(base.size)
        return Image.blend(base, other, 0.5)

    def _create_summary_pdf(self, temp_dir: str, diff_entries: List[Dict], old_file: str, new_file: str) -> str:
        report_doc = fitz.open()
        page = report_doc.new_page()
        y = 72
        line_height = 14
        bottom_limit = fitz.paper_size("letter")[1] - 72

        lines = [
            "Document Comparison Report",
            "",
            f"Old Document: {old_file}",
            f"New Document: {new_file}",
            f"Total Differences: {self.statistics['TOTAL_DIFFERENCES']}",
            f"Deleted Segments: {self.statistics['DELETED_COUNT']}",
            f"Added Segments: {self.statistics['ADDED_COUNT']}",
            "",
            "Structured Diff Summary:",
            "",
        ]

        for index, item in enumerate(diff_entries, start=1):
            lines.append(f"[{index}] 原文档描述: {item['old_desc']}")
            lines.append(f"    原文档页数: {item['old_page']}")
            lines.append(f"    新文档描述: {item['new_desc']}")
            lines.append(f"    新文档页数: {item['new_page']}")
            lines.append("")

        for line in lines:
            if y > bottom_limit:
                page = report_doc.new_page()
                y = 72
            page.insert_text((72, y), line, fontsize=10, fontname="helv")
            y += line_height

        report_file = path.join(temp_dir, "diff_summary.pdf")
        report_doc.save(report_file)
        report_doc.close()
        return report_file

    def _resolve_output_dir(self, source_file: str) -> str:
        if self.OUTPUT_PATH is None:
            return path.dirname(source_file)
        if self.OUTPUT_PATH == "\\":
            return path.dirname(source_file)
        output_dir = self.OUTPUT_PATH.rstrip("\\/")
        if path.isdir(output_dir):
            return output_dir
        return path.dirname(source_file)

    def _apply_output_format(self, image: Image.Image) -> Image.Image:
        if self.OUTPUT_GS:
            image = image.convert("L")
        elif self.OUTPUT_BW:
            image = image.convert("1")
        else:
            image = image.convert("RGB")
        return image

    def handle_files(self, files: List[str]) -> str:
        self.logMessage.emit(f"Processing files:\n    {files[0]}\n    {files[1]}")

        if self.MAIN_PAGE == "New Document":
            new_index, old_index = 0, 1
            main_index = 0
        else:
            new_index, old_index = 1, 0
            main_index = 1

        with fitz.open(files[old_index]) as old_doc, fitz.open(files[new_index]) as new_doc:
            self.statistics["MAIN_PAGE"] = files[main_index]
            total_pages = max(old_doc.page_count, new_doc.page_count)
            self.statistics["NUM_PAGES"] = total_pages

            self.logMessage.emit("Extracting text tokens from old document...")
            old_tokens = self._extract_tokens(old_doc)
            self.progressUpdated.emit(10)

            self.logMessage.emit("Extracting text tokens from new document...")
            new_tokens = self._extract_tokens(new_doc)
            self.progressUpdated.emit(20)

            self.logMessage.emit("Running semantic text diff...")
            diff_entries = self._build_diff_entries(old_tokens, new_tokens)

            old_highlights: Dict[int, List[fitz.Rect]] = {}
            new_highlights: Dict[int, List[fitz.Rect]] = {}
            page_change_counts: Dict[int, int] = {}

            for entry in diff_entries:
                if entry["type"] in ("delete", "replace"):
                    self.statistics["DELETED_COUNT"] += 1
                    for page_idx, rects in entry["old_rects"].items():
                        old_highlights.setdefault(page_idx, []).extend(rects)
                        page_change_counts[page_idx + 1] = page_change_counts.get(page_idx + 1, 0) + len(rects)

                if entry["type"] in ("add", "replace"):
                    self.statistics["ADDED_COUNT"] += 1
                    for page_idx, rects in entry["new_rects"].items():
                        new_highlights.setdefault(page_idx, []).extend(rects)
                        page_change_counts[page_idx + 1] = page_change_counts.get(page_idx + 1, 0) + len(rects)

            self.statistics["TOTAL_DIFFERENCES"] = len(diff_entries)
            self.statistics["PAGES_WITH_DIFFERENCES"] = sorted(page_change_counts.items(), key=lambda item: item[0])
            self.logMessage.emit(f"Semantic diff complete. Found {len(diff_entries)} structured differences.")
            self.progressUpdated.emit(30)

            main_filename = path.basename(files[main_index])
            output_dir = self._resolve_output_dir(files[main_index])
            progress_per_page = 60.0 / max(total_pages, 1)
            current_progress = 30.0
            toc = []

            with TemporaryDirectory() as temp_dir:
                page_artifacts = []

                for page_index in range(total_pages):
                    self.logMessage.emit(f"Rendering page {page_index + 1} / {total_pages}...")
                    old_base, old_page_rect = self._render_page(old_doc, page_index, self.DPI_LEVEL)
                    new_base, new_page_rect = self._render_page(new_doc, page_index, self.DPI_LEVEL)

                    # 动态判定当前页面是否为扫描页，从而按用户设定渲染不同扩展度的选框
                    old_page = old_doc.load_page(page_index) if page_index < old_doc.page_count else None
                    new_page = new_doc.load_page(page_index) if page_index < new_doc.page_count else None
                    
                    old_is_scanned = False
                    if old_page:
                        old_is_scanned = (sum(1 for w in old_page.get_text("words") if (w[4] or "").strip()) < 12) or self.FORCE_OCR
                        
                    new_is_scanned = False
                    if new_page:
                        new_is_scanned = (sum(1 for w in new_page.get_text("words") if (w[4] or "").strip()) < 12) or self.FORCE_OCR

                    old_padding = self.OCR_BOX_PADDING if old_is_scanned else self.VECTOR_BOX_PADDING
                    new_padding = self.OCR_BOX_PADDING if new_is_scanned else self.VECTOR_BOX_PADDING

                    old_marked = self._draw_rectangles(old_base.copy(), old_page_rect, old_highlights.get(page_index, []), (220, 38, 38), old_padding)
                    new_marked = self._draw_rectangles(new_base.copy(), new_page_rect, new_highlights.get(page_index, []), (22, 163, 74), new_padding)

                    output_images = []
                    if self.INCLUDE_IMAGES.get("New Copy", False):
                        output_images.append(("New Copy", self._resize_if_needed(new_marked)))
                    if self.INCLUDE_IMAGES.get("Old Copy", False):
                        output_images.append(("Old Copy", self._resize_if_needed(old_marked)))
                    if self.INCLUDE_IMAGES.get("Markup", True):
                        output_images.append(
                            (
                                "Markup",
                                self._resize_if_needed(new_marked if self.MAIN_PAGE == "New Document" else old_marked),
                            )
                        )
                    if self.INCLUDE_IMAGES.get("Difference", False):
                        output_images.append(("Difference", self._combine_side_by_side(old_marked, new_marked)))
                    if self.INCLUDE_IMAGES.get("Overlay", False):
                        output_images.append(("Overlay", self._overlay_blend(old_marked, new_marked)))

                    if not output_images:
                        output_images.append(("Markup", self._resize_if_needed(new_marked if self.MAIN_PAGE == "New Document" else old_marked)))

                    for variant_index, (label, image) in enumerate(output_images):
                        image = self._apply_output_format(image)
                        image_file = path.join(temp_dir, f"{page_index}_{variant_index}.pdf")
                        image.save(image_file, resolution=self.DPI_LEVEL, author="MAXFIELD", optimize=self.REDUCE_FILESIZE)
                        page_artifacts.append(image_file)
                        toc.append([1, f"Page {page_index + 1} {label}", len(page_artifacts)])

                    current_progress += progress_per_page
                    self.progressUpdated.emit(int(current_progress))

                self.logMessage.emit("Generating structured diff report page...")
                report_file = self._create_summary_pdf(temp_dir, diff_entries, files[old_index], files[new_index])
                page_artifacts.append(report_file)
                toc.append([1, "Structured Diff Summary", len(page_artifacts)])

                self.logMessage.emit("Compiling output PDF...")
                compiled_pdf = fitz.open()
                for pdf_file in page_artifacts:
                    part = fitz.open(pdf_file)
                    compiled_pdf.insert_pdf(part, links=False)
                    part.close()

                compiled_pdf.set_toc(toc)
                output_path = path.join(output_dir, f"{path.splitext(main_filename)[0]} Comparison.pdf")
                output_iterator = 0
                while path.exists(output_path):
                    output_iterator += 1
                    output_path = path.join(
                        output_dir,
                        f"{path.splitext(main_filename)[0]} Comparison Rev {output_iterator}.pdf",
                    )

                compiled_pdf.save(output_path)
                compiled_pdf.close()

        self.progressUpdated.emit(100)
        self.logMessage.emit(f"Comparison file created: {output_path}")
        self.compareComplete.emit(2)
        return output_path


# ---------------------------------------------------------------------------
# 集成主窗口组件 (物理分屏、旋转预览、微调滑块及比对调用)
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Comparison Tool")
        self.resize(550, 720)  # 高度自适应拉伸以完美容纳微调模块
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)

        self.title_bar = CustomTitleBar(self)
        self.title_bar.setObjectName("TitleBar")
        self.setMenuWidget(self.title_bar)

        self.settings = load_settings()
        self.files = [None, None]
        self.compare_thread: Optional["CompareThread"] = None
        self.progress_window: Optional["ProgressWindow"] = None

        # 初始化分栏拖拽区域
        self.drop_label_old = DragDropLabel("old", self)
        self.drop_label_new = DragDropLabel("new", self)

        # 初始化预览容器
        self.preview_old = QLabel()
        self.preview_old.setFixedSize(220, 220)
        self.preview_old.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_old.setStyleSheet("""
            background-color: #FFFFFF;
            border: 2px dashed #FF6B6B;
            border-radius: 8px;
            color: #FF6B6B;
            font: bold 12px "Segoe UI", sans-serif;
        """)

        self.preview_new = QLabel()
        self.preview_new.setFixedSize(220, 220)
        self.preview_new.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_new.setStyleSheet("""
            background-color: #FFFFFF;
            border: 2px dashed #2196F3;
            border-radius: 8px;
            color: #2196F3;
            font: bold 12px "Segoe UI", sans-serif;
        """)

        # 旋转调节按键
        old_btn_style = self._btn_style("#FF6B6B", "#FFF5F5", "#FFE0E0")
        self.rot_ccw_old_btn = QPushButton("↺ 90° CCW")
        self.rot_ccw_old_btn.setStyleSheet(old_btn_style)
        self.rot_ccw_old_btn.clicked.connect(lambda: self.rotate_file_in_place("old", -90))

        self.rot_cw_old_btn = QPushButton("90° CW ↻")
        self.rot_cw_old_btn.setStyleSheet(old_btn_style)
        self.rot_cw_old_btn.clicked.connect(lambda: self.rotate_file_in_place("old", 90))

        new_btn_style = self._btn_style("#2196F3", "#F0F8FF", "#E3F2FD")
        self.rot_ccw_new_btn = QPushButton("↺ 90° CCW")
        self.rot_ccw_new_btn.setStyleSheet(new_btn_style)
        self.rot_ccw_new_btn.clicked.connect(lambda: self.rotate_file_in_place("new", -90))

        self.rot_cw_new_btn = QPushButton("90° CW ↻")
        self.rot_cw_new_btn.setStyleSheet(new_btn_style)
        self.rot_cw_new_btn.clicked.connect(lambda: self.rotate_file_in_place("new", 90))

        # 交换按键
        self.swap_button = QPushButton("⇅ Swap Old ⇄ New")
        self.swap_button.setStyleSheet("""
            QPushButton {
                background-color: #E8F0FE;
                color: #2196F3;
                border: 1px solid #BBDEFB;
                border-radius: 8px;
                padding: 8px;
                font: bold 12px "Segoe UI", Arial, sans-serif;
            }
            QPushButton:hover {
                background-color: #BBDEFB;
            }
        """)
        self.swap_button.clicked.connect(self.swap_files)

        self.compare_button = QPushButton("Compare", self)
        self.compare_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
                font: bold 14px "Segoe UI", Arial, sans-serif;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.compare_button.clicked.connect(self.compare)

        # -------------------------------------------------------------------
        # 新增标框与 OCR 空间精度微调滑块布局 (精简直接嵌入 UI 界面)
        # -------------------------------------------------------------------
        tuning_group = QGroupBox("Drawing & OCR Fine-Tuning (标框与OCR微调)")
        tuning_layout = QVBoxLayout(tuning_group)
        tuning_layout.setSpacing(6)
        tuning_layout.setContentsMargins(12, 10, 12, 12)

        self.vec_pad_label = QLabel(f"Vector Box Expansion (矢量图框扩展): {self.settings.get('VECTOR_BOX_PADDING', 2)} px")
        self.vec_pad_slider = QSlider(Qt.Orientation.Horizontal)
        self.vec_pad_slider.setRange(0, 20)
        self.vec_pad_slider.setValue(self.settings.get("VECTOR_BOX_PADDING", 2))
        self.vec_pad_slider.valueChanged.connect(self.update_vec_pad)

        self.ocr_pad_label = QLabel(f"OCR Box Expansion (图片/OCR框扩展): {self.settings.get('OCR_BOX_PADDING', 2)} px")
        self.ocr_pad_slider = QSlider(Qt.Orientation.Horizontal)
        self.ocr_pad_slider.setRange(0, 20)
        self.ocr_pad_slider.setValue(self.settings.get("OCR_BOX_PADDING", 2))
        self.ocr_pad_slider.valueChanged.connect(self.update_ocr_pad)

        self.ocr_merge_label = QLabel(f"OCR Merge Threshold (OCR合并精度): {self.settings.get('OCR_MERGE_DIST_H', 15)} px")
        self.ocr_merge_slider = QSlider(Qt.Orientation.Horizontal)
        self.ocr_merge_slider.setRange(5, 40)
        self.ocr_merge_slider.setValue(self.settings.get("OCR_MERGE_DIST_H", 15))
        self.ocr_merge_slider.valueChanged.connect(self.update_ocr_merge)

        tuning_layout.addWidget(self.vec_pad_label)
        tuning_layout.addWidget(self.vec_pad_slider)
        tuning_layout.addWidget(self.ocr_pad_label)
        tuning_layout.addWidget(self.ocr_pad_slider)
        tuning_layout.addWidget(self.ocr_merge_label)
        tuning_layout.addWidget(self.ocr_merge_slider)

        self.dpi_label = QLabel("DPI:", self)
        self.dpi_label.setAlignment(Qt.AlignmentFlag.AlignBottom)
        self.dpi_combo = QComboBox(self)
        self.dpi_combo.addItems(self.settings["DPI_LABELS"])
        self.dpi_combo.setCurrentText(self.settings["DPI"])
        self.dpi_combo.currentTextChanged.connect(self.update_dpi)

        self.page_label = QLabel("Page Size:", self)
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignBottom)
        self.page_combo = QComboBox(self)
        self.page_combo.addItems(list(self.settings["PAGE_SIZES"].keys()))
        self.page_combo.setCurrentText(self.settings["PAGE_SIZE"])
        self.page_combo.currentTextChanged.connect(self.update_page_size)

        # 构建紧凑型双栏布局
        central_widget = QWidget()
        new_layout = QVBoxLayout(central_widget)
        new_layout.setContentsMargins(16, 16, 16, 16)
        new_layout.setSpacing(12)

        cols_layout = QHBoxLayout()
        cols_layout.setSpacing(16)

        # 左栏 (Old Document)
        old_col = QVBoxLayout()
        old_col.setSpacing(8)
        self.drop_label_old.setMaximumHeight(50)
        self.drop_label_old.setMinimumHeight(45)
        old_col.addWidget(self.drop_label_old)
        
        preview_old_container = QHBoxLayout()
        preview_old_container.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_old_container.addWidget(self.preview_old)
        old_col.addLayout(preview_old_container)

        old_rot_layout = QHBoxLayout()
        old_rot_layout.setSpacing(6)
        old_rot_layout.addWidget(self.rot_ccw_old_btn)
        old_rot_layout.addWidget(self.rot_cw_old_btn)
        old_col.addLayout(old_rot_layout)

        # 右栏 (New Document)
        new_col = QVBoxLayout()
        new_col.setSpacing(8)
        self.drop_label_new.setMaximumHeight(50)
        self.drop_label_new.setMinimumHeight(45)
        new_col.addWidget(self.drop_label_new)

        preview_new_container = QHBoxLayout()
        preview_new_container.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_new_container.addWidget(self.preview_new)
        new_col.addLayout(preview_new_container)

        new_rot_layout = QHBoxLayout()
        new_rot_layout.setSpacing(6)
        new_rot_layout.addWidget(self.rot_ccw_new_btn)
        new_rot_layout.addWidget(self.rot_cw_new_btn)
        new_col.addLayout(new_rot_layout)

        cols_layout.addLayout(old_col)
        cols_layout.addLayout(new_col)
        new_layout.addLayout(cols_layout)

        new_layout.addWidget(self.swap_button)
        new_layout.addWidget(tuning_group)
        new_layout.addWidget(self.compare_button)

        settings_row = QHBoxLayout()
        dpi_vbox = QVBoxLayout()
        dpi_vbox.setSpacing(4)
        dpi_vbox.addWidget(self.dpi_label)
        dpi_vbox.addWidget(self.dpi_combo)
        
        page_vbox = QVBoxLayout()
        page_vbox.setSpacing(4)
        page_vbox.addWidget(self.page_label)
        page_vbox.addWidget(self.page_combo)
        
        settings_row.addLayout(dpi_vbox)
        settings_row.addLayout(page_vbox)
        new_layout.addLayout(settings_row)

        self.setCentralWidget(central_widget)
        self.set_stylesheet()
        self.update_ui_state()

    def _btn_style(self, color, bg, hover_bg) -> str:
        return f"""
            QPushButton {{
                background-color: {bg};
                color: {color};
                border: 1px solid {color};
                border-radius: 6px;
                padding: 6px;
                font: bold 12px "Segoe UI", sans-serif;
            }}
            QPushButton:hover {{
                background-color: {hover_bg};
            }}
            QPushButton:disabled {{
                color: #BDBDBD;
                background-color: #F5F5F5;
                border-color: #E0E0E0;
            }}
        """

    def set_stylesheet(self):
        self.setStyleSheet("""
            QLabel {
                font: 14px "Segoe UI", Arial, sans-serif;
                color: #1A1A2E;
            }
            QMainWindow {
                background-color: #F0F4F8;
            }
            #TitleBar {
                background-color: #2196F3;
            }
            #SettingsButton {
                background-color: #FFC107;
                color: #1A1A2E;
                border-radius: 6px;
                font-weight: bold;
            }
            #MinimizeButton {
                background-color: transparent;
                color: white;
                font-weight: bold;
            }
            #CloseButton {
                background-color: transparent;
                color: white;
                font-weight: bold;
            }
            #MinimizeButton:hover {
                background-color: rgba(255,255,255,0.2);
            }
            #CloseButton:hover {
                background-color: #FF5252;
            }
            QComboBox {
                height: 32px;
                border-radius: 8px;
                background-color: white;
                border: 1px solid #E0E6ED;
                color: #1A1A2E;
                padding: 4px 8px;
                font: 13px "Segoe UI", Arial, sans-serif;
            }
            QComboBox:hover {
                border: 1px solid #2196F3;
            }
            QComboBox QAbstractItemView {
                padding: 6px;
                background-color: white;
                selection-background-color: #E3F2FD;
                selection-color: #1565C0;
                color: #1A1A2E;
            }
            QComboBox::drop-down {
                border: none;
            }
            QSlider::groove:horizontal {
                border: 1px solid #E0E6ED;
                height: 6px;
                background: #F0F4F8;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                background: #2196F3;
                border: 1px solid #2196F3;
                width: 14px;
                height: 14px;
                margin: -4px 0;
                border-radius: 7px;
            }
            QSlider::handle:horizontal:hover {
                background: #1976D2;
                border-color: #1976D2;
            }
            QGroupBox {
                color: #1A1A2E;
                font: bold 12px "Segoe UI", Arial, sans-serif;
                border: 1px solid #E0E6ED;
                border-radius: 8px;
                margin-top: 6px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px 0 4px;
                color: #2196F3;
            }
        """)

    def update_vec_pad(self, val):
        self.settings["VECTOR_BOX_PADDING"] = val
        self.vec_pad_label.setText(f"Vector Box Expansion (矢量图框扩展): {val} px")
        save_settings(self.settings)

    def update_ocr_pad(self, val):
        self.settings["OCR_BOX_PADDING"] = val
        self.ocr_pad_label.setText(f"OCR Box Expansion (图片/OCR框扩展): {val} px")
        save_settings(self.settings)

    def update_ocr_merge(self, val):
        self.settings["OCR_MERGE_DIST_H"] = val
        self.ocr_merge_label.setText(f"OCR Merge Threshold (OCR合并精度): {val} px")
        save_settings(self.settings)

    def refresh_preview(self, role):
        label = self.drop_label_new if role == "new" else self.drop_label_old
        preview_widget = self.preview_new if role == "new" else self.preview_old
        file_path = label.file_path

        if not file_path or not os.path.exists(file_path):
            preview_widget.clear()
            text = "Drop NEW PDF here\nor click to browse" if role == "new" else "Drop OLD PDF here\nor click to browse"
            preview_widget.setText(text)
            return

        try:
            doc = fitz.open(file_path)
            page = doc.load_page(0)
            
            rect = page.rect
            scale = 210 / max(rect.width, rect.height, 1)
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, dpi=150)
            
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            img.thumbnail((210, 210))
            doc.close()

            data = img.tobytes("raw", "RGB")
            qt_img = QImage(data, img.width, img.height, img.width * 3, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qt_img)
            preview_widget.setPixmap(pixmap)
        except Exception as e:
            preview_widget.setText(f"Preview Error:\n{e}")

    def rotate_file_in_place(self, role, angle_diff):
        label = self.drop_label_new if role == "new" else self.drop_label_old
        file_path = label.file_path
        if not file_path or not os.path.exists(file_path):
            return

        try:
            doc = fitz.open(file_path)
            for page in doc:
                page.set_rotation((page.rotation + angle_diff) % 360)
            
            temp_fd, temp_path = tempfile.mkstemp(suffix=".pdf")
            os.close(temp_fd)
            doc.save(temp_path)
            doc.close()

            shutil.move(temp_path, file_path)
            self.refresh_preview(role)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save modified PDF:\n{e}")

    def update_ui_state(self):
        has_old = bool(self.drop_label_old.file_path and os.path.exists(self.drop_label_old.file_path))
        has_new = bool(self.drop_label_new.file_path and os.path.exists(self.drop_label_new.file_path))

        self.rot_ccw_old_btn.setEnabled(has_old)
        self.rot_cw_old_btn.setEnabled(has_old)
        self.rot_ccw_new_btn.setEnabled(has_new)
        self.rot_cw_new_btn.setEnabled(has_new)

        self.refresh_preview("old")
        self.refresh_preview("new")

    def swap_files(self):
        if self.files and self.files[0] and self.files[1]:
            self.files = [self.files[1], self.files[0]]
            self.drop_label_old.file_path = self.files[1]
            self.drop_label_new.file_path = self.files[0]
            self.drop_label_old._update_style()
            self.drop_label_new._update_style()
            self.update_ui_state()

    def update_dpi(self, dpi):
        if dpi:
            self.settings["DPI"] = dpi
            self.settings["DPI_LEVEL"] = self.settings["DPI_LEVELS"][self.settings["DPI_LABELS"].index(dpi)]
            save_settings(self.settings)

    def update_page_size(self, page_size):
        self.settings["PAGE_SIZE"] = page_size
        save_settings(self.settings)

    def compare(self):
        if self.files and len(self.files) == 2 and self.files[0] and self.files[1]:
            self.progress_window = ProgressWindow()
            self.progress_window.show()
            self.compare_thread = CompareThread(self.files, self.progress_window, self)
            self.compare_thread.finished.connect(self._thread_cleanup)
            self.compare_thread.start()

    def _thread_cleanup(self):
        self.compare_thread = None


# ---------------------------------------------------------------------------
# 全局基本样式配置与程序启动入口
# ---------------------------------------------------------------------------
stylesheet = """
#SettingsButton {
    background-color: #FFC107;
    color: #1A1A2E;
    border-radius: 6px;
    font-weight: bold;
}
#SettingsButton:hover {
    background-color: #FFB300;
}
#MinimizeButton:hover {
    background-color: rgba(255,255,255,0.2);
}
#CloseButton:hover {
    background-color: #FF5252;
}
"""

if __name__ == "__main__":
    app = QApplication([])
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setStyleSheet(stylesheet)
    
    window = MainWindow()
    window.show()
    app.exec()