import re
import sys
from difflib import SequenceMatcher
from json import dump, load
from os import path
from tempfile import TemporaryDirectory
from time import sleep
from typing import Dict, List, Optional, Tuple

import fitz
from PIL import Image, ImageDraw
from PySide6.QtCore import QThread, Signal, Slot, Qt
from PySide6.QtGui import QIcon
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
)


class AdvancedSettings(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = load_settings()

        self.min_diff_label = QLabel("Minimum Diff Token Length [Default: 2]:")
        self.min_diff_desc = QLabel(
            "Short differences are usually noise in PDFs. Increase this value to suppress tiny token-level changes."
        )
        self.min_diff_desc.setWordWrap(True)
        self.min_diff_desc.setStyleSheet("color: black; font: 12px Arial, sans-serif;")
        self.min_diff_spinbox = QSpinBox(self)
        self.min_diff_spinbox.setMinimum(1)
        self.min_diff_spinbox.setMaximum(20)
        self.min_diff_spinbox.setValue(self.settings.get("TEXT_MIN_DIFF_LENGTH", 2))
        self.min_diff_spinbox.valueChanged.connect(self.update_min_diff)

        self.normalize_checkbox = QCheckBox("Normalize Text (lowercase, trim punctuation)")
        self.normalize_checkbox.setChecked(self.settings.get("NORMALIZE_TEXT", True))
        self.normalize_checkbox.stateChanged.connect(self.update_normalize)

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.addWidget(self.min_diff_label)
        layout.addWidget(self.min_diff_desc)
        layout.addWidget(self.min_diff_spinbox)
        layout.addWidget(self.normalize_checkbox)
        self.setLayout(layout)

        self.setStyleSheet(
            """
            QLabel {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QCheckBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
        """
        )

    def update_min_diff(self, value):
        self.settings["TEXT_MIN_DIFF_LENGTH"] = int(value)
        save_settings(self.settings)

    def update_normalize(self, state):
        self.settings["NORMALIZE_TEXT"] = state == 2
        save_settings(self.settings)


class DPISettings(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = window
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

        self.setStyleSheet(
            """
            QLabel {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QComboBox {
                height: 30px;
                border-radius: 5px;
                background-color: #454545;
                selection-background-color: #ff5e0e;
                color: white;
            }
            QComboBox QAbstractItemView {
                padding: 10px;
                background-color: #454545;
                selection-background-color: #ff5e0e;
                color: white;
            }
        """
        )

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

        self.parent_window.dpi_combo.clear()
        self.parent_window.dpi_combo.addItems(self.settings["DPI_LABELS"])
        self.parent_window.dpi_combo.setCurrentText(self.settings["DPI_LABELS"][3])
        save_settings(self.settings)


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

        self.setStyleSheet(
            """
            QLabel {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QComboBox {
                height: 30px;
                border-radius: 5px;
                background-color: #454545;
                selection-background-color: #ff5e0e;
                color: white;
            }
            QComboBox QAbstractItemView {
                padding: 10px;
                background-color: #454545;
                selection-background-color: #ff5e0e;
                color: white;
                font: 14px Arial, sans-serif;
            }
            QCheckBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
        """
        )

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
        self.setFixedSize(500, 500)

        self.tab_widget = QTabWidget(self)
        self.tab_widget.addTab(OutputSettings(), "Output")
        self.tab_widget.addTab(DPISettings(), "DPI")
        self.tab_widget.addTab(AdvancedSettings(), "Advanced")

        layout = QVBoxLayout(self)
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)


class CustomTitleBar(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setFixedHeight(40)
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(10, 0, 10, 0)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        spacer_item = QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.title_label = QLabel("PyPDFCompare")
        self.settings_button = QPushButton("Settings", self)
        self.settings_button.setObjectName("SettingsButton")
        self.settings_button.setFixedSize(65, 25)
        self.settings_button.clicked.connect(self.open_settings)

        self.minimize_button = QPushButton("-", self)
        self.minimize_button.setObjectName("MinimizeButton")
        self.minimize_button.setFixedSize(20, 20)
        self.minimize_button.clicked.connect(self.parent.showMinimized)

        self.close_button = QPushButton("X", self)
        self.close_button.setObjectName("CloseButton")
        self.close_button.setFixedSize(20, 20)
        self.close_button.clicked.connect(self.parent.close)

        self.layout.addWidget(self.settings_button)
        self.layout.addItem(spacer_item)
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
    def __init__(self, parent=None):
        super().__init__(parent)
        self._parent = parent
        self.setAcceptDrops(True)
        size_policy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.clicked.connect(self.browse_files)
        self.setSizePolicy(size_policy)
        self.setStyleSheet(
            """
            QPushButton {
                color: black;
                background-color: #f7f7f7;
                border-radius: 10px;
                border: 2px solid #ff5e0e;
            }
        """
        )
        self.setText("Drop files here or click to browse")

    def browse_files(self) -> None:
        files = list(QFileDialog.getOpenFileNames(self, "Open Files", "", "PDF Files (*.pdf)")[0])
        if files and len(files) == 2:
            self.setText(f"Main File: {files[0]}\nSecondary File: {files[1]}")
            self._parent.files = files

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        files.reverse()
        if files and len(files) == 2:
            self.setText(f"Main File: {files[0]}\nSecondary File: {files[1]}")
            self._parent.files = files


class ProgressWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyPDFCompare")
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
                background-color: #2b2b2b;
            }
            QTextBrowser {
                background-color: #323232;
                color: #c8c8c8;
                border: 1px solid #ff5e0e;
                border-radius: 5px;
            }
            QProgressBar {
                border: 1px solid #ff5e0e;
                border-radius: 5px;
                text-align: center;
                color: #c8c8c8;
                background-color: #202020;
            }
            QProgressBar::chunk {
                background-color: #0075d5;
                width: 1px;
                border: 1px solid transparent;
                border-radius: 5px;
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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Maxfield Auto Markup")
        self.setGeometry(100, 100, 500, 300)
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)

        self.title_bar = CustomTitleBar(self)
        self.title_bar.setObjectName("TitleBar")
        self.setMenuWidget(self.title_bar)

        self.settings = load_settings()
        self.files = None
        self.compare_thread: Optional["CompareThread"] = None
        self.progress_window: Optional["ProgressWindow"] = None

        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        self.drop_label = DragDropLabel(self)
        self.compare_button = QPushButton("Compare", self)
        self.compare_button.clicked.connect(self.compare)

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

        layout.addWidget(self.drop_label)
        layout.addWidget(self.compare_button)
        layout.addWidget(self.dpi_label)
        layout.addWidget(self.dpi_combo)
        layout.addWidget(self.page_label)
        layout.addWidget(self.page_combo)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        self.set_stylesheet()

    def set_stylesheet(self):
        self.drop_label.setStyleSheet(
            """
            QPushButton {
                color: white;
                background-color: #2D2D2D;
                border-radius: 10px;
                border: 2px solid #ff5e0e;
            }
        """
        )

        self.setStyleSheet(
            """
            QLabel {
                font: 14px Arial, sans-serif;
                color: white;
            }
            QMainWindow {
                background-color: #2D2D2D;
            }
            #TitleBar {
                background-color: #1f1f1f;
            }
            #SettingsButton {
                background-color: #ff5e0e;
                color: black;
            }
            #MinimizeButton {
                background-color: #2b2b2b;
            }
            #CloseButton {
                background-color: #2b2b2b;
            }
            #MinimizeButton:hover {
                background-color: blue;
            }
            #CloseButton:hover {
                background-color: red;
            }
        """
        )

    def update_dpi(self, dpi):
        if dpi:
            self.settings["DPI"] = dpi
            self.settings["DPI_LEVEL"] = self.settings["DPI_LEVELS"][self.settings["DPI_LABELS"].index(dpi)]
            save_settings(self.settings)

    def update_page_size(self, page_size):
        self.settings["PAGE_SIZE"] = page_size
        save_settings(self.settings)

    def compare(self):
        if self.files and len(self.files) == 2:
            self.progress_window = ProgressWindow()
            self.progress_window.show()
            self.compare_thread = CompareThread(self.files, self.progress_window, self)
            self.compare_thread.finished.connect(self._thread_cleanup)
            self.compare_thread.start()

    def _thread_cleanup(self):
        self.compare_thread = None


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
        text = text.strip()
        if self.NORMALIZE_TEXT:
            text = text.lower()
            text = re.sub(r"[\s\t\r\n]+", "", text)
            text = re.sub(r"[\.,;:()\[\]{}<>\-_=+`~\"']+", "", text)
        return text

    def _extract_tokens(self, doc: fitz.Document) -> List[Dict]:
        tokens = []
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            words = page.get_text("words")
            words.sort(key=lambda word: (word[5], word[6], word[7], word[1], word[0]))
            for word in words:
                raw = (word[4] or "").strip()
                if not raw:
                    continue
                norm = self._normalize_text(raw)
                if len(norm) < self.TEXT_MIN_DIFF_LENGTH:
                    continue
                tokens.append(
                    {
                        "text": raw,
                        "norm": norm,
                        "page": page_num,
                        "rect": fitz.Rect(word[0], word[1], word[2], word[3]),
                    }
                )
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
    def _draw_rectangles(image: Image.Image, page_rect: fitz.Rect, rects: List[fitz.Rect], color: Tuple[int, int, int]):
        if not rects:
            return image

        draw = ImageDraw.Draw(image)
        x_scale = image.width / max(page_rect.width, 1)
        y_scale = image.height / max(page_rect.height, 1)
        stroke = max(2, int(min(image.width, image.height) / 350))

        for rect in rects:
            x0 = max(0, int(rect.x0 * x_scale) - 2)
            y0 = max(0, int(rect.y0 * y_scale) - 2)
            x1 = min(image.width - 1, int(rect.x1 * x_scale) + 2)
            y1 = min(image.height - 1, int(rect.y1 * y_scale) + 2)
            draw.rectangle((x0, y0, x1, y1), outline=color, width=stroke)

        return image

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

                    old_marked = self._draw_rectangles(old_base.copy(), old_page_rect, old_highlights.get(page_index, []), (220, 38, 38))
                    new_marked = self._draw_rectangles(new_base.copy(), new_page_rect, new_highlights.get(page_index, []), (22, 163, 74))

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


def save_settings(settings: dict) -> None:
    settings_path = "settings.json"
    with open(settings_path, "w", encoding="utf-8") as file:
        dump(settings, file, indent=4)


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
            "Low DPI: Viewing Only [150]",
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


stylesheet = """
#SettingsButton {
    background-color: #ff5e0e;
    color: black;
}
#MinimizeButton:hover {
    background-color: blue;
}
#CloseButton:hover {
    background-color: red;
}
#SettingsDialog {
    color: black;
}
"""


if __name__ == "__main__":
    app = QApplication([])
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setStyleSheet(stylesheet)
    window = MainWindow()
    window.show()
    app.exec()