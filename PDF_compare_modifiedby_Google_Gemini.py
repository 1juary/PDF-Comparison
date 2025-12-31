import fitz  # PyMuPDF
from pathlib import Path
# from PyInstaller.utils.hooks import collect_submodules # 如果需要打包exe可取消注释
from typing import List, Tuple, Optional
from os import path
from time import sleep
from json import load, dump
from tempfile import TemporaryDirectory

# 引入数值计算和图像处理库
import numpy as np
import cv2
from PIL import Image, ImageChops, ImageDraw, ImageOps, ImageFont

from PySide6.QtCore import QThread, Signal, Slot, Qt
from PySide6.QtWidgets import QMainWindow, QProgressBar, QApplication, QWidget, QVBoxLayout, QTextBrowser, QDialog, \
    QFrame, QPushButton, QLabel, \
    QSpinBox, QDoubleSpinBox, QComboBox, QCheckBox, QLineEdit, QGroupBox, QTabWidget, QStyleFactory, QFormLayout, \
    QHBoxLayout, QSpacerItem, QSizePolicy, QFileDialog, QGridLayout
from PySide6.QtGui import QIcon


# =================================================================================
# GUI 组件 (保持原样)
# =================================================================================

class AdvancedSettings(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = load_settings()

        self.threshold_label = QLabel("Threshold [Default: 128]:")
        self.threshold_desc = QLabel(
            "To analyze the pdf, it must be thresholded (converted to pure black and white). The threshold setting controls the point at which pixels become white or black determined based upon grayscale color values of 0-255")
        self.threshold_desc.setWordWrap(True)
        self.threshold_desc.setStyleSheet("color: white; font: 12px Arial, sans-serif;")
        self.threshold_spinbox = QSpinBox(self)
        self.threshold_spinbox.setMinimum(0)
        self.threshold_spinbox.setMaximum(255)
        self.threshold_spinbox.setValue(self.settings["THRESHOLD"])
        self.threshold_spinbox.valueChanged.connect(self.update_threshold)

        self.minimum_area_label = QLabel("Minimum Area [Default: 100]:")
        self.minimum_area_desc = QLabel(
            "When marking up the pdf, boxes are created to highlight major changes. The minimum area setting controls the minimum size the boxes can be which will ultimately control what becomes classfied as a significant change.")
        self.minimum_area_desc.setWordWrap(True)
        self.minimum_area_desc.setStyleSheet("color: white; font: 12px Arial, sans-serif;")
        self.minimum_area_spinbox = QSpinBox(self)
        self.minimum_area_spinbox.setMinimum(0)
        self.minimum_area_spinbox.setMaximum(1000)
        self.minimum_area_spinbox.setValue(self.settings["MIN_AREA"])
        self.minimum_area_spinbox.valueChanged.connect(self.update_area)

        self.epsilon_label = QLabel("Precision [Default: 0.00]:")
        self.epsilon_desc = QLabel(
            "When marking up the pdf, outlines are created to show any change. The precision setting controls the maximum distance of the created contours around a change. Smaller values will have better precision and follow curves better and higher values will have more space between the contour and the change.")
        self.epsilon_desc.setWordWrap(True)
        self.epsilon_desc.setStyleSheet("color: white; font: 12px Arial, sans-serif;")
        self.epsilon_spinbox = QDoubleSpinBox(self)
        self.epsilon_spinbox.setMinimum(0.000)
        self.epsilon_spinbox.setMaximum(1.000)
        self.epsilon_spinbox.setSingleStep(0.001)
        self.epsilon_spinbox.setValue(self.settings["EPSILON"])
        self.epsilon_spinbox.valueChanged.connect(self.update_epsilon)

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        layout.addWidget(self.epsilon_label)
        layout.addWidget(self.epsilon_desc)
        layout.addWidget(self.epsilon_spinbox)
        layout.addWidget(self.minimum_area_label)
        layout.addWidget(self.minimum_area_desc)
        layout.addWidget(self.minimum_area_spinbox)
        layout.addWidget(self.threshold_label)
        layout.addWidget(self.threshold_desc)
        layout.addWidget(self.threshold_spinbox)

        self.setLayout(layout)
        self.setStyleSheet('''   
            QLabel {
                color: white;
                font: 14px Arial, sans-serif;
            }
            QSpinBox, QDoubleSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
                background-color: #f0f0f0;
            }
        ''')

    def update_threshold(self, threshold):
        self.settings["THRESHOLD"] = threshold
        save_settings(self.settings)

    def update_area(self, area):
        self.settings["MIN_AREA"] = area
        save_settings(self.settings)

    def update_epsilon(self, epsilon):
        self.settings["EPSILON"] = epsilon
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
        self.setStyleSheet('''
            QLabel {
                color: white;
                font: 14px Arial, sans-serif;
            }
            QSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
                background-color: #f0f0f0;
            }
        ''')

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
            self.settings["OUTPUT_PATH"] if self.output_path_combobox.currentText() == "Specified Path" else "")
        self.specified_entry.textChanged.connect(self.set_output_path)

        self.checkbox_image1 = QCheckBox("New Copy")
        if self.settings["INCLUDE_IMAGES"]["New Copy"] is True:
            self.checkbox_image1.setChecked(True)
        self.checkbox_image2 = QCheckBox("Old Copy")
        if self.settings["INCLUDE_IMAGES"]["Old Copy"] is True:
            self.checkbox_image2.setChecked(True)
        self.checkbox_image3 = QCheckBox("Markup")
        if self.settings["INCLUDE_IMAGES"]["Markup"] is True:
            self.checkbox_image3.setChecked(True)
        self.checkbox_image4 = QCheckBox("Difference")
        if self.settings["INCLUDE_IMAGES"]["Difference"] is True:
            self.checkbox_image4.setChecked(True)
        self.checkbox_image5 = QCheckBox("Overlay")
        if self.settings["INCLUDE_IMAGES"]["Overlay"] is True:
            self.checkbox_image5.setChecked(True)

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
        self.setStyleSheet('''
            QLabel {
                color: white;
                font: 14px Arial, sans-serif;
            }
            QSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QLineEdit {
                color: black;
                background-color: #f0f0f0;
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
                color: white;
                font: 14px Arial, sans-serif;
            }
            QGroupBox {
                color: white;
                font: bold 14px Arial, sans-serif;
                border: 1px solid silver;
                border-radius: 6px;
                margin-top: 6px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 7px;
                padding: 0 3px 0 3px;
            }
        ''')

    def set_output_path(self, option):
        if option == "Source Path":
            self.settings["OUTPUT_PATH"] = None
        elif option == "Default Path":
            self.settings["OUTPUT_PATH"] = "\\"
        else:
            self.settings["OUTPUT_PATH"] = self.specified_entry.text()
            self.settings["OUTPUT_PATH"].replace("\\", "\\\\")
            if not self.settings["OUTPUT_PATH"].endswith("\\"):
                self.settings["OUTPUT_PATH"] += "\\"

        # Don't save if it's triggered by text change and combox is not "Specified"
        # but logic here is simple update
        save_settings(self.settings)

    def set_output_images(self, state):
        checkbox = self.sender()
        if state == 2:
            self.settings["INCLUDE_IMAGES"][checkbox.text()] = True
        else:
            self.settings["INCLUDE_IMAGES"][checkbox.text()] = False
        save_settings(self.settings)

    def set_scaling(self, state):
        if state == 2:
            self.settings["SCALE_OUTPUT"] = True
        else:
            self.settings["SCALE_OUTPUT"] = False
        save_settings(self.settings)

    def set_bw(self, state):
        if state == 2:
            self.settings["OUTPUT_BW"] = True
        else:
            self.settings["OUTPUT_BW"] = False
        save_settings(self.settings)

    def set_gs(self, state):
        if state == 2:
            self.settings["OUTPUT_GS"] = True
        else:
            self.settings["OUTPUT_GS"] = False
        save_settings(self.settings)

    def set_reduced_filesize(self, state):
        if state == 2:
            self.settings["REDUCE_FILESIZE"] = True
        else:
            self.settings["REDUCE_FILESIZE"] = False
        save_settings(self.settings)

    def set_main_page(self, page):
        self.settings["MAIN_PAGE"] = page
        save_settings(self.settings)


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setWindowModality(Qt.ApplicationModal)
        self.setFixedSize(500, 600)

        self.tab_widget = QTabWidget(self)

        output_settings = OutputSettings()
        dpi_settings = DPISettings()
        advanced_settings = AdvancedSettings()

        self.tab_widget.addTab(output_settings, "Output")
        self.tab_widget.addTab(dpi_settings, "DPI")
        self.tab_widget.addTab(advanced_settings, "Advanced")

        layout = QVBoxLayout(self)
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)
        self.setStyleSheet('''
            QDialog {
                background-color: #333333;
                color: white;
            }
            QTabWidget::pane {
                border: 1px solid #444;
                background: #333;
            }
            QTabBar::tab {
                background: #555;
                color: white;
                padding: 8px 20px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: #ff5e0e;
                color: black;
            }
        ''')


class CustomTitleBar(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setFixedHeight(40)
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(10, 0, 10, 0)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        spacer_item = QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.title_label = QLabel("PyPDFCompare (Region Enhanced)")
        self.settings_button = QPushButton("Settings", self)
        self.settings_button.setObjectName('SettingsButton')
        self.settings_button.setFixedSize(65, 25)
        self.settings_button.clicked.connect(self.open_settings)
        self.minimize_button = QPushButton("-", self)
        self.minimize_button.setObjectName('MinimizeButton')
        self.minimize_button.setFixedSize(20, 20)
        self.minimize_button.clicked.connect(self.parent.showMinimized)
        self.close_button = QPushButton("X", self)
        self.close_button.setObjectName('CloseButton')
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
        if self.draggable:
            if event.button() == Qt.LeftButton:
                self.drag_start_position = event.globalPosition().toPoint() - self.parent.pos()
        event.accept()

    def mouseMoveEvent(self, event):
        if self.draggable and self.drag_start_position is not None:
            if event.buttons() == Qt.LeftButton:
                if (
                        event.globalPosition().toPoint() - self.drag_start_position
                ).manhattanLength() > self.dragging_threshold:
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
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.clicked.connect(self.browse_files)
        self.setSizePolicy(sizePolicy)
        self.setStyleSheet("""
        QPushButton {
                color: black;
                background-color: #f7f7f7;
                border-radius: 10px;
                border: 2px solid #ff5e0e;
            }
        """)
        self.setText("Drop files here or click to browse")

    def browse_files(self) -> None:
        files = list(QFileDialog.getOpenFileNames(self, "Open Files", "", "PDF Files (*.pdf)")[0])
        self.update_parent_files(files)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        files.reverse()
        self.update_parent_files(files)

    def update_parent_files(self, files):
        if files and len(files) == 2:
            self._parent.files = files
            # 假设 Files[0] 是新文件(New/Blue), Files[1] 是旧文件(Old/Red)
            # 在CompareThread里：files[0] if self.MAIN_PAGE == "New Document"
            self._parent.file_new_edit.setText(files[0])
            self._parent.file_old_edit.setText(files[1])
            self.setText(f"Files Selected")


class ProgressWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyPDFCompare")
        self.resize(600, 500)

        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)
        self.layout = QVBoxLayout()

        self.progressBar = QProgressBar()
        self.logArea = QTextBrowser()
        self.logArea.setReadOnly(True)

        self.layout.addWidget(self.progressBar)
        self.layout.addWidget(self.logArea)

        self.centralWidget.setLayout(self.layout)

        self.setStyleSheet("""
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
        """)

    @Slot(int)
    def update_progress(self, progress):
        self.progressBar.setValue(progress)

    @Slot(str)
    def update_log(self, message):
        self.logArea.append(message)

    @Slot(int)
    def operation_complete(self, time):
        sleep(time)
        self.close()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Maxfield Auto Markup")
        self.setGeometry(100, 100, 500, 350)
        # self.setWindowIcon(QIcon("app_icon.png")) # 找不到图标时会报错，建议注释或确保文件存在
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.title_bar = CustomTitleBar(self)
        self.title_bar.setObjectName("TitleBar")
        self.setMenuWidget(self.title_bar)
        self.settings = load_settings()
        self.files = None

        layout = QVBoxLayout()

        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # 1. 修改：拖放区域只作为操作入口
        self.drop_label = DragDropLabel(self)
        layout.addWidget(self.drop_label)

        # 2. 新增：独立显示文件路径的区域 + 交换按钮
        file_group = QGroupBox("Selected Documents")

        file_display_layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.file_old_edit = QLineEdit()
        self.file_old_edit.setReadOnly(True)
        self.file_old_edit.setPlaceholderText("Reference File (Old) - Red Markup")
        self.file_old_edit.setStyleSheet("color: #ff5555; font-weight: bold;")

        self.file_new_edit = QLineEdit()
        self.file_new_edit.setReadOnly(True)
        self.file_new_edit.setPlaceholderText("New File (Current) - Blue Markup")
        self.file_new_edit.setStyleSheet("color: #55aaff; font-weight: bold;")

        form_layout.addRow(QLabel("Old (Ref):"), self.file_old_edit)
        form_layout.addRow(QLabel("New (Cur):"), self.file_new_edit)

        self.swap_button = QPushButton("⇅ Swap New/Old Files")
        self.swap_button.setStyleSheet("""
            QPushButton {
                background-color: #444; 
                color: white; 
                border: 1px solid #777;
                border-radius: 4px;
                padding: 4px;
            }
            QPushButton:hover {
                background-color: #555;
            }
        """)
        self.swap_button.clicked.connect(self.swap_files)

        file_display_layout.addLayout(form_layout)
        file_display_layout.addWidget(self.swap_button)

        file_group.setLayout(file_display_layout)
        layout.addWidget(file_group)

        # 按钮和其他设置
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
        self.drop_label.setStyleSheet("""
        QPushButton {
            color: white;
            background-color: #2D2D2D;
            border-radius: 10px;
            border: 2px solid #ff5e0e;
            }""")

        self.setStyleSheet("""
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
            QGroupBox {
                color: white;
                font-weight: bold;
                border: 1px solid #555;
                margin-top: 6px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)

    def swap_files(self):
        """交换新旧文件的槽函数"""
        if self.files and len(self.files) == 2:
            # 交换列表中的元素
            self.files = [self.files[1], self.files[0]]
            # 更新UI显示 (注意 update_parent_files 逻辑是 files[0] -> New, files[1] -> Old)
            self.file_new_edit.setText(self.files[0])
            self.file_old_edit.setText(self.files[1])

    def update_dpi(self, dpi):
        if dpi != "":
            self.settings["DPI"] = dpi
            self.settings["DPI_LEVEL"] = self.settings["DPI_LEVELS"][self.settings["DPI_LABELS"].index(dpi)]
            save_settings(self.settings)

    def update_page_size(self, page_size):
        self.settings["PAGE"] = page_size
        self.settings["PAGE_SIZE"] = self.settings["PAGE_SIZES"][page_size]
        save_settings(self.settings)

    def compare(self):
        if self.files and len(self.files) == 2:
            progress_window = ProgressWindow()
            progress_window.show()
            compare_thread = CompareThread(self.files, progress_window, self)
            compare_thread.start()


# =================================================================================
# 核心修改区域：CompareThread (区域特征增强版)
# =================================================================================

class CompareThread(QThread):
    progressUpdated = Signal(int)
    compareComplete = Signal(int)
    logMessage = Signal(str)

    def __init__(self, files: List[str], progress_window: ProgressWindow, parent=None):
        super(CompareThread, self).__init__(parent)
        compare_settings = load_settings()
        self.DPI_LEVEL = compare_settings.get("DPI_LEVEL")
        self.PAGE_SIZE = tuple(compare_settings.get("PAGE_SIZES").get(compare_settings.get("PAGE_SIZE")))
        self.INCLUDE_IMAGES = compare_settings.get("INCLUDE_IMAGES")
        self.MAIN_PAGE = compare_settings.get("MAIN_PAGE")
        self.THRESHOLD = compare_settings.get("THRESHOLD")

        self.STRUCTURAL_MERGE_KERNEL = max(3, int(self.DPI_LEVEL / 25))

        self.MIN_AREA = compare_settings.get("MIN_AREA")
        self.EPSILON = compare_settings.get("EPSILON")
        self.OUTPUT_PATH = compare_settings.get("OUTPUT_PATH")
        self.SCALE_OUTPUT = compare_settings.get("SCALE_OUTPUT")
        self.OUTPUT_BW = compare_settings.get("OUTPUT_BW")
        self.OUTPUT_GS = compare_settings.get("OUTPUT_GS")
        self.REDUCE_FILESIZE = compare_settings.get("REDUCE_FILESIZE")
        self.files = files
        self.progress_window = progress_window
        self.statistics = {
            "NUM_PAGES": 0,
            "MAIN_PAGE": None,
            "TOTAL_DIFFERENCES": 0,
            "PAGES_WITH_DIFFERENCES": []
        }

        self.progressUpdated.connect(self.progress_window.update_progress)
        self.logMessage.connect(self.progress_window.update_log)
        self.compareComplete.connect(self.progress_window.operation_complete)

    def run(self):
        try:
            self.handle_files(self.files)
        except Exception as e:
            self.logMessage.emit(f"Error during processing: {str(e)}")
            import traceback
            traceback.print_exc()

    def align_images_global(self, img1_cv, img2_cv):
        """ 全局对齐，用于修正整体的旋转或位移 """
        gray1 = cv2.cvtColor(img1_cv, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2_cv, cv2.COLOR_BGR2GRAY)

        max_features = 5000
        orb = cv2.ORB_create(max_features)
        keypoints1, descriptors1 = orb.detectAndCompute(gray1, None)
        keypoints2, descriptors2 = orb.detectAndCompute(gray2, None)

        if descriptors1 is None or descriptors2 is None:
            return img2_cv

        matcher = cv2.DescriptorMatcher_create(cv2.DESCRIPTOR_MATCHER_BRUTEFORCE_HAMMING)
        matches = matcher.match(descriptors1, descriptors2, None)
        matches.sort(key=lambda x: x.distance, reverse=False)
        num_good_matches = int(len(matches) * 0.15)
        matches = matches[:num_good_matches]

        if len(matches) < 10:
            return img2_cv

        points1 = np.zeros((len(matches), 2), dtype=np.float32)
        points2 = np.zeros((len(matches), 2), dtype=np.float32)

        for i, match in enumerate(matches):
            points1[i, :] = keypoints1[match.queryIdx].pt
            points2[i, :] = keypoints2[match.trainIdx].pt

        h, mask = cv2.findHomography(points2, points1, cv2.RANSAC)
        if h is None:
            return img2_cv

        height, width, channels = img1_cv.shape
        aligned_img = cv2.warpPerspective(img2_cv, h, (width, height))
        return aligned_img

    def get_content_regions(self, img_bgr) -> List[Tuple[int, int, int, int]]:
        """
        基于形态学，分离出非白色背景的独立区域（ROI）
        返回: List of (x, y, w, h)
        """
        gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)

        # 1. 二值化 (反转，内容变白，背景变黑)
        # 假设背景是白色，阈值250以上算背景
        _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)

        # 2. 形态学膨胀 (连接相邻文字成块)
        # 横向膨胀力度大一些，纵向小一些，适合文本行
        # 根据 DPI 动态调整核大小
        h_k = max(5, int(self.DPI_LEVEL / 20))
        v_k = max(2, int(self.DPI_LEVEL / 60))
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_k, v_k))

        dilated = cv2.dilate(thresh, kernel, iterations=3)

        # 3. 查找轮廓
        contours, hierarchy = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        regions = []
        min_w = int(self.DPI_LEVEL / 10)  # 过滤太小的噪点
        min_h = int(self.DPI_LEVEL / 10)

        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            if w > min_w and h > min_h:
                regions.append((x, y, w, h))

        # 从上到下排序
        regions.sort(key=lambda r: r[1])
        return regions

    def add_visualization_legend(self, image: Image.Image) -> Image.Image:
        """
        在图片顶部添加图例说明
        [ 蓝色框: 新增 ] [ 红色框: 删除 ] [ 橙色框: 修改 ]
        """
        w, h = image.size
        # 增加高度以容纳图例
        legend_height = max(60, int(h * 0.05))

        # 创建一个白条放在顶部
        new_img = Image.new("RGB", (w, h + legend_height), (255, 255, 255))
        new_img.paste(image, (0, legend_height))

        draw = ImageDraw.Draw(new_img)

        # 动态字体大小
        font_size = max(12, int(w / 80))
        # 尝试加载默认字体，如果不可用则忽略
        try:
            # Linux/Windows通用尝试
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()

        # 定义图例项
        items = [
            ("New/Added", "blue", (200, 220, 255)),
            ("Old/Removed", "red", (255, 200, 200)),
            ("Modified/Typos", "orange", None)
        ]

        # 1. 计算每个图例项的宽度
        box_size = int(font_size * 1.5)
        padding = 10
        item_spacing = 30

        item_widths = []
        for text, color, fill in items:
            try:
                # Pillow >= 9.2.0
                text_w = draw.textlength(text, font=font)
            except:
                # Fallback estimate
                text_w = len(text) * font_size * 0.6

            total_item_w = box_size + padding + text_w
            item_widths.append(total_item_w)

        # 2. 计算总宽度并确定起始位置 (居中)
        total_legend_width = sum(item_widths) + (len(items) - 1) * item_spacing

        start_x = (w - total_legend_width) // 2
        # 如果图例太宽超出页面，则靠左对齐 (至少保留10px边距)
        if start_x < 10:
            start_x = 10

        y_center = legend_height // 2
        box_top = y_center - box_size // 2

        current_x = start_x

        # 3. 绘制
        for i, (text, color, fill_color) in enumerate(items):
            # Draw Box
            if fill_color:
                draw.rectangle([current_x, box_top, current_x + box_size, box_top + box_size], outline=color,
                               fill=fill_color, width=2)
            else:
                draw.rectangle([current_x, box_top, current_x + box_size, box_top + box_size], outline=color, width=2)

            # Draw Text
            draw.text((current_x + box_size + padding, box_top - 2), text, fill="black", font=font)

            # Move to next
            current_x += item_widths[i] + item_spacing

        return new_img

    def mark_differences(self, page_num: int, image1_pil: Image.Image, image2_pil: Image.Image) -> List[Image.Image]:
        """
        区域特征对比算法 (Region-based Feature Comparison)
        """

        # 转 OpenCV 格式
        img1_cv = cv2.cvtColor(np.array(image1_pil), cv2.COLOR_RGB2BGR)
        img2_cv = cv2.cvtColor(np.array(image2_pil), cv2.COLOR_RGB2BGR)

        # 0. 全局粗对其 (Pre-alignment)
        if img1_cv.shape == img2_cv.shape:
            try:
                img2_cv = self.align_images_global(img1_cv, img2_cv)
            except:
                pass

        # 准备画布
        diff_canvas = np.ones_like(img1_cv) * 255  # 白底差异图
        overlay_canvas_r = np.array(image1_pil.convert("L"))  # 红色通道(新内容/差异)
        overlay_canvas_gb = np.array(image1_pil.convert("L"))  # 青色通道(旧内容/基准)

        # 标记用的图层
        markup_layer = Image.new('RGBA', image1_pil.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(markup_layer)

        # 1. 提取 Page 1 (基准) 的内容区域
        regions1 = self.get_content_regions(img1_cv)

        # 用于记录 Page 2 哪些像素已经被匹配过了
        page2_matched_mask = np.zeros(img2_cv.shape[:2], dtype=np.uint8)

        gray1 = cv2.cvtColor(img1_cv, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2_cv, cv2.COLOR_BGR2GRAY)

        total_differences = 0

        # 2. 遍历 Page 1 的区域，去 Page 2 找朋友
        for (x1, y1, w1, h1) in regions1:
            # 提取 ROI 模板
            template = gray1[y1:y1 + h1, x1:x1 + w1]

            # 定义搜索范围 (Search Window)
            # 假设内容可能上下移动，但水平移动不会太大
            # 垂直搜索范围：上下各 30% 页面高度
            search_h_margin = int(img1_cv.shape[0] * 0.3)
            search_w_margin = int(img1_cv.shape[1] * 0.1)  # 水平 10%

            y_search_start = max(0, y1 - search_h_margin)
            y_search_end = min(img1_cv.shape[0], y1 + h1 + search_h_margin)
            x_search_start = max(0, x1 - search_w_margin)
            x_search_end = min(img1_cv.shape[1], x1 + w1 + search_w_margin)

            # 搜索区域
            search_region = gray2[y_search_start:y_search_end, x_search_start:x_search_end]

            match_found = False
            best_match_rect = None  # (x, y, w, h) in Page 2 global coords

            # 只有当搜索区域比模板大时才能搜索
            if search_region.shape[0] > h1 and search_region.shape[1] > w1:
                res = cv2.matchTemplate(search_region, template, cv2.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

                # 相似度阈值 (0.6 比较宽松，因为可能有文字修改)
                if max_val > 0.6:
                    match_found = True
                    # 计算在 Page 2 的全局坐标
                    match_x = x_search_start + max_loc[0]
                    match_y = y_search_start + max_loc[1]
                    best_match_rect = (match_x, match_y, w1, h1)

            if match_found and best_match_rect:
                mx, my, mw, mh = best_match_rect

                # 标记该区域已在 Page 2 被认领
                cv2.rectangle(page2_matched_mask, (mx, my), (mx + mw, my + mh), 255, -1)

                # --- 局部精确对比 ---
                roi2 = gray2[my:my + mh, mx:mx + mw]

                # 计算差异
                local_diff = cv2.absdiff(template, roi2)
                _, local_thresh = cv2.threshold(local_diff, self.THRESHOLD, 255, cv2.THRESH_BINARY)

                # 统计差异像素
                diff_pixels = cv2.countNonZero(local_thresh)

                # 如果差异较多，说明内容有修改 (Typo / Modification)
                if diff_pixels > 50:  # 噪点过滤
                    # 在差异图上画出来
                    # Page 1 位置画红色 (旧内容)
                    diff_canvas[y1:y1 + h1, x1:x1 + w1][local_thresh > 0] = [0, 0, 255]
                    # Page 2 位置画蓝色 (新内容) - 可选，为了不混乱，我们通常画在 Page 1 对应位置
                    # 或者我们可以将 Page 2 的差异也映射回 Page 1 的坐标系显示

                    # 在 Markup 层画黄框 (表示位置找到了，但有修改)
                    # 查找具体的差异轮廓
                    l_cnts, _ = cv2.findContours(local_thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                    for lc in l_cnts:
                        if cv2.contourArea(lc) > 20:  # 忽略微小噪点
                            lx, ly, lw, lh = cv2.boundingRect(lc)
                            # 映射回 Page 1 坐标
                            draw.rectangle([x1 + lx, y1 + ly, x1 + lx + lw, y1 + ly + lh], outline="orange", width=2)
                            total_differences += 1

                # 更新 Overlay 画布 (将 Page 2 的匹配内容搬运到 Page 1 的位置来做叠图)
                # 这样即使段落错行，叠图也是重合的！
                overlay_canvas_r[y1:y1 + h1, x1:x1 + w1] = roi2

            else:
                # 没找到匹配 -> 这是一个被删除的段落 (Deleted Block)
                # 在 Markup 画红框
                draw.rectangle([x1, y1, x1 + w1, y1 + h1], outline="red", width=3, fill=(255, 0, 0, 50))
                diff_canvas[y1:y1 + h1, x1:x1 + w1] = [0, 0, 255]  # 涂红
                total_differences += 1

        # 3. 检查 Page 2 中“未被访问”的区域 -> 新增内容 (Inserted Content)
        # 获取 Page 2 的所有内容区域
        regions2 = self.get_content_regions(img2_cv)

        for (x2, y2, w2, h2) in regions2:
            # 检查这个区域的中心点是否已经被 mask 覆盖
            cx, cy = x2 + w2 // 2, y2 + h2 // 2
            if page2_matched_mask[cy, cx] == 0:
                # 这是一个新增区域
                # 在 Markup 画蓝框
                # 注意：这是 Page 2 的坐标，我们需要画在 Page 1 的图上
                # 如果是纯新增，可能覆盖在 Page 1 的空白处，或者与其他内容重叠
                draw.rectangle([x2, y2, x2 + w2, y2 + h2], outline="blue", width=3, fill=(0, 0, 255, 50))

                # 在差异图上画蓝
                # 注意边界检查
                if y2 + h2 < diff_canvas.shape[0] and x2 + w2 < diff_canvas.shape[1]:
                    # 获取该区域的 mask
                    roi_gray2 = gray2[y2:y2 + h2, x2:x2 + w2]
                    _, roi_mask = cv2.threshold(roi_gray2, 250, 255, cv2.THRESH_BINARY_INV)
                    diff_canvas[y2:y2 + h2, x2:x2 + w2][roi_mask > 0] = [255, 0, 0]

                # 在 Overlay 的红色通道显示新增内容
                if y2 + h2 < overlay_canvas_r.shape[0] and x2 + w2 < overlay_canvas_r.shape[1]:
                    # 这里比较暴力，直接覆盖，可能会遮挡 Page 1 的内容
                    # 但既然是新增的，那个位置在 Page 1 通常是空白
                    pass
                    # 为了 Overlay 效果，我们其实不需要动 overlay_canvas_r 对应位置，
                    # 因为它初始化就是 Page 1。
                    # 我们需要让这部分变红。
                    # Overlay 逻辑： R通道=Page2(Aligned), GB通道=Page1
                    # 对于新增块，Page2 有字(黑)，Page1 无字(白) -> R=0, GB=255 -> Cyan (青色)
                    # 等等，之前的逻辑是红青互补。
                    # 我们需要把 Page 2 的这个新增块，贴到 overlay_canvas_r 上
                    overlay_canvas_r[y2:y2 + h2, x2:x2 + w2] = gray2[y2:y2 + h2, x2:x2 + w2]

                total_differences += 1

        # 4. 生成最终图片
        markup_image = Image.alpha_composite(image1_pil.convert("RGBA"), markup_layer)
        markup_image = self.add_visualization_legend(markup_image)  # Add Legend

        # Overlay 合成: R=Page2(Reconstructed), GB=Page1
        overlay_image = Image.merge("RGB", (
        Image.fromarray(overlay_canvas_r), image1_pil.convert("L"), image1_pil.convert("L")))

        diff_image = Image.fromarray(cv2.cvtColor(diff_canvas.astype(np.uint8), cv2.COLOR_BGR2RGB))

        # 统计
        if total_differences > 0:
            self.statistics["TOTAL_DIFFERENCES"] += total_differences
            self.statistics["PAGES_WITH_DIFFERENCES"].append((page_num, total_differences))

        # 输出打包
        output = []
        target_size = (int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL))

        def resize_if_needed(img):
            if self.SCALE_OUTPUT:
                return img.resize(target_size)
            return img

        if self.INCLUDE_IMAGES["New Copy"]:
            output.append(resize_if_needed(image1_pil if self.MAIN_PAGE == "New Document" else image2_pil))
        if self.INCLUDE_IMAGES["Old Copy"]:
            output.append(resize_if_needed(image2_pil if self.MAIN_PAGE == "New Document" else image1_pil))
        if self.INCLUDE_IMAGES["Markup"]:
            output.append(resize_if_needed(markup_image.convert("RGB")))
        if self.INCLUDE_IMAGES["Difference"]:
            output.append(resize_if_needed(diff_image))
        if self.INCLUDE_IMAGES["Overlay"]:
            output.append(resize_if_needed(overlay_image))

        return output

    def pdf_to_image(self, page_number: int, doc: fitz.Document) -> Image.Image:
        if page_number < doc.page_count:
            pix = doc.load_page(page_number).get_pixmap(dpi=self.DPI_LEVEL)
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if pix.n == 4:  # RGBA -> RGB
                img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
            image = Image.fromarray(img_array)
        else:
            ref_page = doc.load_page(0)
            rect = ref_page.rect
            width = int(rect.width * self.DPI_LEVEL / 72)
            height = int(rect.height * self.DPI_LEVEL / 72)
            image = Image.new("RGB", (width, height), (255, 255, 255))

        if self.SCALE_OUTPUT:
            image = image.resize((int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL)))
        return image

    def handle_files(self, files: List[str]) -> str:
        self.logMessage.emit(f"Processing files:\n    {files[0]}\n    {files[1]}")
        toc = []
        current_progress = 0

        with fitz.open(files[0 if self.MAIN_PAGE == "New Document" else 1]) as doc1, fitz.open(
                files[0 if self.MAIN_PAGE == "OLD" else 1]) as doc2:

            size = doc1.load_page(0).rect
            if self.PAGE_SIZE[0] is None:
                self.PAGE_SIZE = (size.width / 72, size.height / 72)
                # 重新计算 kernel
                self.STRUCTURAL_MERGE_KERNEL = max(3, int(self.DPI_LEVEL / 25))

            self.statistics["MAIN_PAGE"] = files[0 if self.MAIN_PAGE == "New Document" else 1]
            filename = files[0 if self.MAIN_PAGE == "New Document" else 1].split("/")[-1]
            source_path = False
            if self.OUTPUT_PATH is None:
                self.OUTPUT_PATH = files[0].replace(filename, "")
                source_path = True

            total_operations = max(doc1.page_count, doc2.page_count)
            self.logMessage.emit(f"Total pages {total_operations}.")
            progress_per_operation = 100.0 / total_operations

            self.logMessage.emit("Creating temporary directory...")
            with TemporaryDirectory() as temp_dir:
                self.logMessage.emit(f"Temporary directory created: {temp_dir}")
                image_files = []

                # Processing Loop
                for i in range(total_operations):
                    self.logMessage.emit(f"Processing page {i + 1} of {total_operations}...")

                    self.logMessage.emit(f"Converting main page...")
                    image1 = self.pdf_to_image(i, doc1)

                    self.logMessage.emit(f"Converting secondary page...")
                    image2 = self.pdf_to_image(i, doc2)

                    self.logMessage.emit(f"Marking differences...")
                    markups = self.mark_differences(i, image1, image2)

                    del image1, image2

                    # Save images
                    self.logMessage.emit(f"Saving output files...")
                    for j, image in enumerate(markups):
                        if self.OUTPUT_GS is True:
                            image = image.convert("L")
                        if self.OUTPUT_BW is True:
                            image = image.convert("1")
                        else:
                            image = image.convert("RGB")

                        image_file = path.join(temp_dir, f"{i}_{j}.pdf")
                        image.save(image_file, resolution=self.DPI_LEVEL, author="MAXFIELD",
                                   optimize=self.REDUCE_FILESIZE)
                        del image
                        image_files.append(image_file)
                        toc.append([1, f"Page {i + 1} Variation {j + 1}", i * len(markups) + j])

                    current_progress += progress_per_operation
                    self.progressUpdated.emit(int(current_progress))

                # Create Stats Page
                text = f"Document Comparison Report\n\nTotal Pages: {total_operations}\nFiles Compared:\n    File in Blue_{files[0]}\n    File in Red_{files[1]}\nMain Page: {self.statistics['MAIN_PAGE']}\nTotal Differences: {self.statistics['TOTAL_DIFFERENCES']}\nPages with differences:\n"
                for page_info in self.statistics["PAGES_WITH_DIFFERENCES"]:
                    text += f"    Page {page_info[0] + 1} Changes: {page_info[1]}\n"

                stats_doc = fitz.open()
                stats_page = stats_doc.new_page()
                text_blocks = text.split('\n')
                y_position = 72
                for line in text_blocks:
                    if y_position > fitz.paper_size('letter')[1] - 72:
                        stats_page = stats_doc.new_page()
                        y_position = 72
                    stats_page.insert_text((72, y_position), line, fontsize=11, fontname="helv")
                    y_position += 12

                stats_filename = path.join(temp_dir, "stats.pdf")
                stats_doc.save(stats_filename)
                stats_doc.close()

                # 插入统计页到最前面 (或按原逻辑处理)
                # 原逻辑似乎是在image_files列表最前面加入了stats_filename，但这里是在循环后生成的
                # 我们把它加到列表最前面
                image_files.insert(0, stats_filename)

                # Compiling PDF
                self.logMessage.emit("Compiling PDF from output folder...")
                compiled_pdf = fitz.open()
                for img_path in image_files:
                    img = fitz.open(img_path)
                    compiled_pdf.insert_pdf(img, links=False)
                    img.close()

                compiled_pdf.set_toc(toc)

                self.logMessage.emit(f"Saving final PDF...")
                output_path = f"{self.OUTPUT_PATH}{filename.split('.')[0]} Comparison.pdf"
                output_iterator = 0

                while path.exists(output_path):
                    output_iterator += 1
                    output_path = f"{self.OUTPUT_PATH}{filename.split('.')[0]} Comparison Rev {output_iterator}.pdf"

                compiled_pdf.save(output_path)
                compiled_pdf.close()

                self.logMessage.emit(f"Comparison file created: {output_path}")
                if source_path:
                    self.OUTPUT_PATH = None

        self.compareComplete.emit(5)
        return output_path


# =================================================================================
# 辅助函数与入口 (保持原样)
# =================================================================================

def save_settings(settings: dict) -> None:
    settings_path = "settings.json"
    if settings_path:
        with open(settings_path, "w") as f:
            dump(settings, f, indent=4)


def load_settings() -> dict:
    settings = None
    settings_path = "settings.json"

    if settings_path and path.exists(settings_path):
        with open(settings_path, "r") as f:
            settings = load(f)
    if not settings:
        settings = _load_default_settings()
    save_settings(settings)
    return settings


def _load_default_settings() -> dict:
    default_settings = {
        "PAGE_SIZES": {
            "AUTO": [None, None],
            "LETTER": [8.5, 11],
            "ANSI A": [11, 8.5],
            "ANSI B": [17, 11],
            "ANSI C": [22, 17],
            "ANSI D": [34, 22]
        },
        "DPI_LEVELS": [75, 150, 300, 600, 1200, 1800],
        "DPI_LABELS": [
            "Low DPI: Draft Quality [75]",
            "Low DPI: Viewing Only [150]",
            "Medium DPI: Printable [300]",
            "Standard DPI [600]",
            "High DPI [1200]: Professional Quality",
            "Max DPI [1800]: Large File Size"
        ],
        "INCLUDE_IMAGES": {"New Copy": False, "Old Copy": False, "Markup": True, "Difference": True, "Overlay": True},
        "DPI": "Standard DPI [600]",
        "DPI_LEVEL": 600,
        "PAGE_SIZE": "AUTO",
        "THRESHOLD": 128,
        "MIN_AREA": 100,
        "EPSILON": 0.0,
        "OUTPUT_PATH": None,
        "SCALE_OUTPUT": True,
        "OUTPUT_BW": False,
        "OUTPUT_GS": False,
        "REDUCE_FILESIZE": True,
        "MAIN_PAGE": "New Document"
    }
    return default_settings


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
