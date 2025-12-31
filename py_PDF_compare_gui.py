import fitz
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules
from typing import List
from os import path
from time import sleep
from json import load, dump
from tempfile import TemporaryDirectory
from numpy import array, where, all, int32
from PIL import Image, ImageChops, ImageDraw, ImageOps
from cv2 import findContours, threshold, approxPolyDP, arcLength, contourArea, boundingRect, THRESH_BINARY, \
    RETR_EXTERNAL, CHAIN_APPROX_SIMPLE

from PySide6.QtCore import QThread, Signal, Slot, Qt
from PySide6.QtWidgets import QMainWindow, QProgressBar, QApplication, QWidget, QVBoxLayout, QTextBrowser, QDialog, \
    QFrame, QPushButton, QLabel, \
    QSpinBox, QDoubleSpinBox, QComboBox, QCheckBox, QLineEdit, QGroupBox, QTabWidget, QStyleFactory, QFormLayout, \
    QHBoxLayout, QSpacerItem, QSizePolicy, QFileDialog
from PySide6.QtGui import QIcon


class AdvancedSettings(QWidget):  #处理高级设置选项的GUI组件,继承自Qwidget
    def __init__(self, parent=None): #唯一指定可选参数，parent，如果不传parent，窗口将作为独立窗口显示
        super().__init__(parent)   #实现特定功能的自定义
        self.settings = load_settings()     #自定义函数加载设置，应用逻辑需要，与QWidget 无关

        self.threshold_label = QLabel("Threshold [Default: 128]:") #自定义UI控件，Qlabel实例，显示一个文本标签，用于说明旁边的 QSpinBox（数值输入框）的作用
        self.threshold_desc = QLabel(
            "To analyze the pdf, it must be thresholded (converted to pure black and white). The threshold setting controls the point at which pixels become white or black determined based upon grayscale color values of 0-255")
        self.threshold_desc.setWordWrap(True) #启用文本自动换行
        self.threshold_desc.setStyleSheet("color: black; font: 12px Arial, sans-serif;")
        self.threshold_spinbox = QSpinBox(self)
        self.threshold_spinbox.setMinimum(0)
        self.threshold_spinbox.setMaximum(255)
        self.threshold_spinbox.setValue(self.settings["THRESHOLD"])
        self.threshold_spinbox.valueChanged.connect(self.update_threshold) #实现交互逻辑，valueChanged 信号会隐式传递一个整数参数（即 threshold）给 update_threshold

        self.minimum_area_label = QLabel("Minimum Area [Default: 100]:")
        self.minimum_area_desc = QLabel(
            "When marking up the pdf, boxes are created to highlight major changes. The minimum area setting controls the minimum size the boxes can be which will ultimately control what becomes classfied as a significant change.")
        self.minimum_area_desc.setWordWrap(True)
        self.minimum_area_desc.setStyleSheet("color: black; font: 12px Arial, sans-serif;")
        self.minimum_area_spinbox = QSpinBox(self)
        self.minimum_area_spinbox.setMinimum(0)
        self.minimum_area_spinbox.setMaximum(1000)
        self.minimum_area_spinbox.setValue(self.settings["MIN_AREA"])
        self.minimum_area_spinbox.valueChanged.connect(self.update_area)

        self.epsilon_label = QLabel("Precision [Default: 0.00]:")
        self.epsilon_desc = QLabel(
            "When marking up the pdf, outlines are created to show any change. The precision setting controls the maximum distance of the created contours around a change. Smaller values will have better precision and follow curves better and higher values will have more space between the contour and the change.")
        self.epsilon_desc.setWordWrap(True)
        self.epsilon_desc.setStyleSheet("color: black; font: 12px Arial, sans-serif;")
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

        self.setLayout(layout) #关键步骤，将具体布局设置到widget
        #setStyleSheet， 美化操作
        self.setStyleSheet('''   
            QLabel {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QSpinBox, QDoubleSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
        ''')

    def update_threshold(self, threshold):#更新阈值设置,valueChanged 默认接收当前spinBox改变的值，信号内部已经绑定了当前值，update_threshold 需要声明接收这个参数：
        self.settings["THESHOLD"] = threshold
        save_settings(self.settings)

    def update_area(self, area): #更新最小区域设置
        self.settings["MIN_AREA"] = area
        save_settings(self.settings)

    def update_epsilon(self, epsilon): #更新精度设置
        self.settings["EPSILON"] = epsilon
        save_settings(self.settings)


class DPISettings(QWidget):  #处理DPI设置选项的GUI组件，六种DPI级别设置
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = window
        self.settings = load_settings()
        self.low_draft_label = QLabel("Low DPI - Draft Quality:")
        self.low_draft_spinbox = QSpinBox(self)
        self.low_draft_spinbox.setMinimum(1)
        self.low_draft_spinbox.setMaximum(99)
        self.low_draft_spinbox.setValue(self.settings["DPI_LEVELS"][0]) #初始化spinbox的值，从seting 文件中提取
        self.low_draft_spinbox.valueChanged.connect(self.update_dpi_levels) #连接信号，spin做出修改之后，对setting 文件进行更新

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
                color: black;
                font: 14px Arial, sans-serif;
            }
            QSpinBox {
                color: black;
                font: 14px Arial, sans-serif;
            }
            QComboBox {  #下拉选择框控件
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
        ''')

    def update_dpi_levels(self, new_dpi): #更新DPI级别设置，一个值，可能同时属于多个层级，需要更新所有相关层级的标签，如果互斥，会导致各个层级DPI的界面显示不一致
        if new_dpi < 100:
            self.settings["DPI_LEVELS"][0] = new_dpi
            self.settings["DPI_LABELS"][0] = f"Low DPI: Draft Quality [{new_dpi}]" #<100,草稿质量
        elif new_dpi < 200:
            self.settings["DPI_LEVELS"][1] = new_dpi
            self.settings["DPI_LABELS"][1] = f"Low DPI: Viewing Quality [{new_dpi}]" #<200,查看质量
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
        save_settings()


class OutputSettings(QWidget): #处理输出设置选项的GUI组件
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = load_settings()
        self.output_path_label = QLabel("Output Path:")
        self.output_path_combobox = QComboBox(self)
        self.output_path_combobox.addItems(["Source Path", "Default Path", "Specified Path"])  #下拉框提供三种路径模式
        if self.settings["OUTPUT_PATH"] == "\\": #输出到默认文件夹
            self.output_path_combobox.setCurrentText("Default Path")
        elif self.settings["OUTPUT_PATH"] is None:
            self.output_path_combobox.setCurrentText("Source Path")
        else:
            self.output_path_combobox.setCurrentText("Specified Path") #自定义路径
        self.output_path_combobox.currentTextChanged.connect(self.set_output_path) #当下拉框选项变化时，调用set_output_path 更新配置

        self.specified_label = QLabel("Specified Path:")
        self.specified_entry = QLineEdit(self) #用于输入自定义路径
        self.specified_entry.setText( #仅当选中Specified Path时显示当前路径
            self.settings["OUTPUT_PATH"] if self.output_path_combobox.currentText() == "Specified Path" else "")
        self.specified_entry.textChanged.connect(self.set_output_path) #文本变化时更新配置

        self.checkbox_image1 = QCheckBox("New Copy") #勾选选项
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

        #每个复选框绑定到 set_output_images
        self.checkbox_image1.stateChanged.connect(self.set_output_images)
        self.checkbox_image2.stateChanged.connect(self.set_output_images)
        self.checkbox_image3.stateChanged.connect(self.set_output_images)
        self.checkbox_image4.stateChanged.connect(self.set_output_images)
        self.checkbox_image5.stateChanged.connect(self.set_output_images)

        #是否缩放页面
        self.scaling_checkbox = QCheckBox("Scale Pages")
        self.scaling_checkbox.setChecked(self.settings["SCALE_OUTPUT"])
        self.scaling_checkbox.stateChanged.connect(self.set_scaling)
        #是否黑白输出
        self.bw_checkbox = QCheckBox("Black/White")
        self.bw_checkbox.setChecked(self.settings["OUTPUT_BW"])
        self.bw_checkbox.stateChanged.connect(self.set_bw)
        #是否灰度输出
        self.gs_checkbox = QCheckBox("Grayscale")
        self.gs_checkbox.setChecked(self.settings["OUTPUT_GS"])
        self.gs_checkbox.stateChanged.connect(self.set_gs)
        #是否减小文件大小
        self.reduce_checkbox = QCheckBox("Reduce Size")
        self.reduce_checkbox.setChecked(self.settings["REDUCE_FILESIZE"])
        self.reduce_checkbox.stateChanged.connect(self.set_reduced_filesize)
        #主页面选择
        self.main_page_label = QLabel("Main Page:")
        self.main_page_combobox = QComboBox(self)
        self.main_page_combobox.addItems(["New Document", "Old Document"])
        self.main_page_combobox.setCurrentText(self.settings["MAIN_PAGE"])
        self.main_page_combobox.currentTextChanged.connect(self.set_main_page)

        #用QGroupBox将相关控件分组（分组创建实例），提升界面可读性
        output_path_group = QGroupBox("Output Settings")
        include_images_group = QGroupBox("Files to include:")
        general_group = QGroupBox("General")
        checkboxes_group = QGroupBox()
        other_group = QGroupBox()

        #输出路径部分使用 QFormLayout（标签+控件对齐）
        output_path_layout = QFormLayout()
        output_path_layout.addRow(self.output_path_label, self.output_path_combobox)
        output_path_layout.addRow(self.specified_label, self.specified_entry)
        output_path_group.setLayout(output_path_layout)

        #文件类型复选框使用水平布局 QHBoxLayout
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

        #所有分组框垂直排列 (QVBoxLayout)
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        main_layout.addWidget(output_path_group)
        main_layout.addWidget(include_images_group)
        main_layout.addWidget(general_group)

        self.setLayout(main_layout)
        self.setStyleSheet('''
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
        ''')

    def set_output_path(self, option):  #由信号传递的当前 QComboBox 选中项的文本（"Source Path"/"Default Path"/"Specified Path"）
        if option == "Source Path":
            self.settings["OUTPUT_PATH"] = None
        elif option == "Default Path":
            self.settings["OUTPUT_PATH"] = "\\" #这里只是用'\\' 做占位符
        else:
            self.settings["OUTPUT_PATH"] = self.specified_entry.text()
            self.settings["OUTPUT_PATH"].replace("\\", "\\\\") #将单反斜杠转义为双反斜杠，确保windows路径合规,为windows特化处理
            self.settings["OUTPUT_PATH"] += "\\" #强制添加末尾分隔符（如将 "C:\output" 转为 "C:\\output\\"）

        save_settings(self.settings)

    def set_output_images(self, state):#设置包含哪些输出图像
        checkbox = self.sender()
        if state == 2:
            self.settings["INCLUDE_IMAGES"][checkbox.text()] = True
        else:
            self.settings["INCLUDE_IMAGES"][checkbox.text()] = False
        save_settings(self.settings)

    def set_scaling(self, state): #设置包含哪些输出图像
        if state == 2:
            self.settings["SCALE_OUTPUT"] = True
        else:
            self.settings["SCALE_OUTPUT"] = False
        save_settings(self.settings)

    def set_bw(self, state): #设置是否黑白输出
        if state == 2:
            self.settings["OUTPUT_BW"] = True
        else:
            self.settings["OUTPUT_BW"] = False
        save_settings(self.settings)

    def set_gs(self, state): #设置是否灰度输出
        if state == 2:
            self.settings["OUTPUT_GS"] = True
        else:
            self.settings["OUTPUT_GS"] = False
        save_settings(self.settings)

    def set_reduced_filesize(self, state): #设置是否减小文件大小
        if state == 2:
            self.settings["REDUCE_FILESIZE"] = True
        else:
            self.settings["REDUCE_FILESIZE"] = False
        save_settings(self.settings)

    def set_main_page(self, page): #设置主页面
        self.settings["MAIN_PAGE"] = page
        save_settings(self.settings)


class SettingsDialog(QDialog): #设置对话框，整合所有设置选项
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setWindowModality(Qt.ApplicationModal)
        self.setFixedSize(500, 500)

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
                color: black;
            }
            
            QLabel {
                color: black;
            }
            
            QSpinBox, QDoubleSpinBox {
                color: black;
            }
        ''')


class CustomTitleBar(QFrame): #自定义标题栏，实现自定义窗口标题栏，替换原生窗口标题栏并提供自定义按钮和拖拽功能。基于 QFrame（而非 QWidget），便于边框样式控制
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent #保存父窗口引用，便于后续操作（如最小化/关闭窗口）。
        self.setFixedHeight(40) #固定标题栏高度为40像素（视觉统一）
        self.layout = QHBoxLayout(self) #水平布局
        self.layout.setContentsMargins(10, 0, 10, 0) #边距左右各10pixel，上下无间隔
        self.layout.setAlignment(Qt.AlignmentFlag.AlignLeft) #默认左对齐（后续通过伸缩项实现左右分栏）

        #将标题文本推到中间，实现左右分栏布局
        spacer_item = QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.title_label = QLabel("PyPDFCompare")  #主程序，显示应用程序名称
        #设置按钮
        self.settings_button = QPushButton("Settings", self)
        self.settings_button.setObjectName('SettingsButton') #设置ObjectName，用于CSS样式定制
        self.settings_button.setFixedSize(65, 25) #固定大小65X25像素
        #和open_settings连接
        self.settings_button.clicked.connect(self.open_settings)
        #窗口最小化按钮
        self.minimize_button = QPushButton("-", self)
        self.minimize_button.setObjectName('MinimizeButton')
        self.minimize_button.setFixedSize(20, 20)
        self.minimize_button.clicked.connect(self.parent.showMinimized) #点击后最小化父窗口
        #窗口关闭按钮
        self.close_button = QPushButton("X", self)
        self.close_button.setObjectName('CloseButton')
        self.close_button.setFixedSize(20, 20)
        self.close_button.clicked.connect(self.parent.close) #点击后关闭父窗口

        #顺序 [设置按钮] + [伸缩项] + [标题] + [伸缩项] + [最小化按钮] + [关闭按钮]
        self.layout.addWidget(self.settings_button)
        self.layout.addItem(spacer_item)
        self.layout.addWidget(self.title_label)
        self.layout.addItem(spacer_item)
        self.layout.addWidget(self.minimize_button)
        self.layout.addWidget(self.close_button)

        #拖拽功能实现
        self.draggable = True #启用拖拽
        self.dragging_threshold = 5 #防误触阈值，
        self.drag_start_position = None #记录拖拽启示坐标

    def mousePressEvent(self, event): #鼠标点击事件
        if self.draggable:
            if event.button() == Qt.LeftButton:
                self.drag_start_position = event.globalPosition().toPoint() - self.parent.pos() #当左键按下时，计算鼠标相对于窗口的偏移量
        event.accept() #保存该偏移量用于后续拖拽计算

    def mouseMoveEvent(self, event):
        if self.draggable and self.drag_start_position is not None:
            if event.buttons() == Qt.LeftButton:
                if (
                        event.globalPosition().toPoint() - self.drag_start_position
                ).manhattanLength() > self.dragging_threshold:  #manhattanLength() 计算移动距离（避免平方根开销），移动超过阈值
                    self.parent.move(event.globalPosition().toPoint() - self.drag_start_position)   #超过阈值后，更新窗口位置为当前鼠标位置 - 初始偏移量
                    self.drag_start_position = event.globalPosition().toPoint() - self.parent.pos() #实时更新drag_start_position,实现平滑拖拽
        event.accept()

    def mouseReleaseEvent(self, event):  #清除拖拽状态
        if event.button() == Qt.LeftButton:
            self.drag_start_position = None
        event.accept()

    def open_settings(self):
        settings_dialog = SettingsDialog(self.parent)
        settings_dialog.exec()  #exec() 阻塞主窗口直到对话框关闭


class DragDropLabel(QPushButton): #拖放区域组件，支持拖放PDF文件
    def __init__(self, parent=None):
        super().__init__(parent)
        self._parent = parent  #保存父窗口引用
        self.setAcceptDrops(True) #启用拖放功能
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding) #水平方向最小宽度，垂直方向可扩展
        self.clicked.connect(self.browse_files) #点击按钮触发文件选择
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

    def browse_files(self) -> None: #打开文件对话框，限制选择PDF文件，可多选文件
        files = list(QFileDialog.getOpenFileNames(self, "Open Files", "", "PDF Files (*.pdf)")[0])
        if files and len(files) == 2:
            self.setText(f"Main File: {files[0]}\nSecondary File: {files[1]}")
            self._parent.files = files #将路径传递给主窗口

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): #检查拖拽内容是否包含文件路径
            event.acceptProposedAction() #接收拖拽操作

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()] #获取文件路径
        files.reverse()  #反转列表以保证顺序
        if files and len(files) == 2:
            self.setText(f"Main File: {files[0]}\nSecondary File: {files[1]}")
            self._parent.files = files


class ProgressWindow(QMainWindow): #进度显示窗口
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
    def update_progress(self, progress): #跟新进度
        self.progressBar.setValue(progress)

    @Slot(str)
    def update_log(self, message): #更新日志
        self.logArea.append(message)

    @Slot(int)
    def operation_complete(self, time): #操作完成处理
        sleep(time)
        self.close()


class MainWindow(QMainWindow): #应用程序主窗口
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Maxfield Auto Markup")
        self.setGeometry(100, 100, 500, 300)
        self.setWindowIcon(QIcon("app_icon.png"))
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        #创建自定义标题栏实例
        self.title_bar = CustomTitleBar(self)
        self.title_bar.setObjectName("TitleBar")
        self.setMenuWidget(self.title_bar)
        #加载用户配置到self.setting字典
        self.settings = load_settings()
        self.files = None

        layout = QVBoxLayout()

        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        #创建支持拖放文件和点击浏览的交互区域
        self.drop_label = DragDropLabel(self)
        #点击时触发compare()方法启动文件比较
        self.compare_button = QPushButton("Compare", self)
        self.compare_button.clicked.connect(self.compare)
        #DPI 下拉框
        self.dpi_label = QLabel("DPI:", self)
        self.dpi_label.setAlignment(Qt.AlignmentFlag.AlignBottom)
        self.dpi_combo = QComboBox(self)
        self.dpi_combo.addItems(self.settings["DPI_LABELS"])
        self.dpi_combo.setCurrentText(self.settings["DPI"])  #设置下拉框的当前选中项，要求参数与下拉框中的某一项的文本完全一致，这部分写在setting.json中
        self.dpi_combo.currentTextChanged.connect(self.update_dpi) #当前选项变化时更新配置
        #页面尺寸下拉框
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
        #布局管理
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.set_stylesheet()

    def set_stylesheet(self):  #设置界面样式
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
        """)

    #更新DPI配置
    def update_dpi(self, dpi):  #更新DPI设置，dpi是由信号自动传递的当前选中项文本，"Standard DPI [600]"，类型是字符串
        if dpi != "":
            self.settings["DPI"] = dpi
            self.settings["DPI_LEVEL"] = self.settings["DPI_LEVELS"][self.settings["DPI_LABELS"].index(dpi)] #在DPI标签列表（如 ["Low DPI [75]", "Standard DPI [600]", ...]）中查找当前选项的索引位置
            save_settings(self.settings)

    def update_page_size(self, page_size): #更新页面大小设置
        self.settings["PAGE"] = page_size
        self.settings["PAGE_SIZE"] = self.settings["PAGE_SIZES"][page_size]
        save_settings(self.settings)

    def compare(self): #启动比较操作
        if self.files and len(self.files) == 2:
            progress_window = ProgressWindow()
            progress_window.show()
            compare_thread = CompareThread(self.files, progress_window, self)
            compare_thread.start()


class CompareThread(QThread): #比较线程，执行核心比较功能
    progressUpdated = Signal(int) #定义信号，传递进度百分比，这个信号会携带一个整数参数
    compareComplete = Signal(int)
    logMessage = Signal(str)

    def __init__(self, files: List[str], progress_window: ProgressWindow, parent=None): #process_windows: ProcessWindow, 进程窗口类
        super(CompareThread, self).__init__(parent) #调用父类的构造函数，等价于super().__init__(parent)
        compare_settings = load_settings() #加载配置
        self.DPI_LEVEL = compare_settings.get("DPI_LEVEL")
        self.PAGE_SIZE = tuple(compare_settings.get("PAGE_SIZES").get(compare_settings.get("PAGE_SIZE")))
        self.INCLUDE_IMAGES = compare_settings.get("INCLUDE_IMAGES")
        self.MAIN_PAGE = compare_settings.get("MAIN_PAGE")
        self.THRESHOLD = compare_settings.get("THRESHOLD")
        self.MERGE_THRESHOLD = int(self.DPI_LEVEL / 100 * self.PAGE_SIZE[0] * self.PAGE_SIZE[1]) if self.PAGE_SIZE[0] and self.PAGE_SIZE[1] else None
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

    def run(self): #线程入口函数
        try:
            self.handle_files(self.files)
        except fitz.fitz.FileDataError as e: #捕获文件数据错误异常
            self.logMessage.emit(f"Error opening file: {e}") #发射日志信号

    def mark_differences(self, page_num: int, image1: Image.Image, image2: Image.Image) -> List[Image.Image]: #标记差异函数
        # Overlay Image
        if self.INCLUDE_IMAGES["Overlay"] is True:
            if image1.size != image2.size:
                image2.size = image2.resize(image1.size)
                self.logMessage.emit("Comparison",   #发射日志信号
                                     "Page sizes don't match and the 'Scale Pages' setting is off, attempting to match page sizes... results may be inaccurate.")
            image1array = array(image1)
            image2array = array(image2)
            image2array[~all(image2array == [255, 255, 255], axis=-1)] = [255, 0,0]  # Convert non-white pixels in image2array to red for overlay.
            #image2array 是一个形状为 (height, width, 3) 的三维数组，最后一个维度是颜色通道，通常使用[R,G,B]的数据表示
            #~取反，得到所有非白色像素的布尔索引。判断每个像素是否为白色，对每个像素的三个通道都判断，得到
            overlay_image = Image.fromarray(
                where(all(image1array == [255, 255, 255], axis=-1, keepdims=True), image2array, image1array)) #如果 image1array 的像素是白色，则取 image2array 的对应像素（已将非白色像素变为红色），否则保留imagearray的原像素
            #all(..., axis=-1, keepdims=True) 会判断每个像素的三个通道是否都为 255（即该像素是白色），结果 shape 为 (height, width, 1)，每个像素只有一个布尔值，表示是否为白色。
            del image1array, image2array 

        # Markup Image / Differences Image
        if self.INCLUDE_IMAGES["Markup"] is True or self.INCLUDE_IMAGES["Difference"] is True:
            diff_image = Image.fromarray(where(all(array(
                ImageOps.colorize(ImageOps.invert(ImageChops.subtract(image2, image1).convert("L")), black="blue",
                                  white="white").convert("RGB")) == [255, 255, 255], axis=-1)[:, :, None], array(
                ImageOps.colorize(ImageOps.invert(ImageChops.subtract(image1, image2).convert("L")), black="red",
                                  white="white").convert("RGB")), array(
                ImageOps.colorize(ImageOps.invert(ImageChops.subtract(image2, image1).convert("L")), black="blue",
                                  white="white").convert("RGB"))))
            if self.INCLUDE_IMAGES["Markup"] is True:
                contours, _ = findContours(
                    threshold(array(ImageChops.difference(image2, image1).convert("L")), self.THRESHOLD, 255,
                              THRESH_BINARY)[1], RETR_EXTERNAL, CHAIN_APPROX_SIMPLE)
                del _
                marked_image = Image.new("RGBA", image1.size, (255, 0, 0, 255)) #创建一个红色背景的透明图层
                marked_image.paste(image1, (0, 0)) #将原图粘贴到透明图层上
                marked_image_draw = ImageDraw.Draw(marked_image) #创建一个可以在图像上绘图的对象
                existing_boxes = []

                for contour in contours:
                    approx = approxPolyDP(contour, (self.EPSILON + 0.0000000001) * arcLength(contour, False), False) #轮廓近似，多变拟合
                    marked_image_draw.line(tuple(map(tuple, array(approx).reshape((-1, 2)).astype(int32))),
                                           fill=(255, 0, 0, 255), width=int(self.DPI_LEVEL / 100)) #在图像上绘制轮廓线

                    if self.MIN_AREA < contourArea(contour):
                        x, y, w, h = boundingRect(contour) #计算轮廓的边界框
                        new_box = (x, y, x + w, y + h) #元组，分别表示左上角和右下角坐标

                        #判断是否与现有框重叠，若重叠则合并
                        merged = False
                        for i, existing_box in enumerate(existing_boxes):
                            # 定义重叠条件，这里使用简单的距离阈值判断
                            if (max(new_box[0], existing_box[0]) - min(new_box[2],
                                                                       existing_box[2]) <= self.MERGE_THRESHOLD and max(
                                new_box[1], existing_box[1]) - min(new_box[3],
                                                                   existing_box[3]) <= self.MERGE_THRESHOLD):
                                #合并框，新框的左上角和右下角坐标取最小值和最大值
                                merged_box = (min(new_box[0], existing_box[0]), min(new_box[1], existing_box[1]),
                                              max(new_box[2], existing_box[2]), max(new_box[3], existing_box[3]))
                                existing_boxes[i] = merged_box  # Update the existing box with the merged one
                                merged = True
                                break #跳出循环，避免重复合并，保证每个框只被合并一次

                        if not merged:
                            existing_boxes.append(new_box)

                # After processing all contours, draw the boxes
                for box in existing_boxes:
                    diff_box = Image.new("RGBA", (box[2] - box[0], box[3] - box[1]), (0, 255, 0, 64))
                    ImageDraw.Draw(diff_box).rectangle([(0, 0), (box[2] - box[0] - 1, box[3] - box[1] - 1)],
                                                       outline=(255, 0, 0, 255), width=int(self.DPI_LEVEL / 100))
                    marked_image.paste(diff_box, (box[0], box[1]), mask=diff_box) #将差异框粘贴到标记图像上，mask表示用diff_box的透明度信息作为掩码,只有diff_box不透明的部分才会覆盖marked_image 
                del contours, marked_image_draw #释放内存
                if len(existing_boxes):
                    self.statistics["TOTAL_DIFFERENCES"] += len(existing_boxes)
                    self.statistics["PAGES_WITH_DIFFERENCES"].append((page_num, len(existing_boxes)))

        # Output
        output = []
        if not self.SCALE_OUTPUT:
            if self.INCLUDE_IMAGES["New Copy"]:
                output.append(image1.resize((int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(
                    self.PAGE_SIZE[1] * self.DPI_LEVEL))) if self.MAIN_PAGE == "New Document" else image2.resize(
                    (int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL))))
            if self.INCLUDE_IMAGES["Old Copy"]:
                output.append(image2.resize((int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(
                    self.PAGE_SIZE[1] * self.DPI_LEVEL))) if self.MAIN_PAGE == "New Document" else image1.resize(
                    (int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL))))
            if self.INCLUDE_IMAGES["Markup"]:
                output.append(marked_image.resize(
                    (int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL))))
            if self.INCLUDE_IMAGES["Difference"]:
                output.append(diff_image.resize(
                    (int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL))))
            if self.INCLUDE_IMAGES["Overlay"]:
                output.append(overlay_image.resize(
                    (int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL))))
        else:
            if self.INCLUDE_IMAGES["New Copy"]:
                output.append(image1 if self.MAIN_PAGE == "New Document" else image2)
            if self.INCLUDE_IMAGES["Old Copy"]:
                output.append(image2 if self.MAIN_PAGE == "New Document" else image1)
            if self.INCLUDE_IMAGES["Markup"]:
                output.append(marked_image)
            if self.INCLUDE_IMAGES["Difference"]:
                output.append(diff_image)
            if self.INCLUDE_IMAGES["Overlay"]:
                output.append(overlay_image)
        return output

    def pdf_to_image(self, page_number: int, doc: fitz.Document) -> Image.Image: #将PDF转换成图像
        if page_number < doc.page_count:
            pix = doc.load_page(page_number).get_pixmap(dpi=self.DPI_LEVEL)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        else:
            pix = doc.load_page(0).get_pixmap(dpi=self.DPI_LEVEL)
            image = Image.new("RGB", (pix.width, pix.height), (255, 255, 255))
        del pix
        if self.SCALE_OUTPUT is True:
            image = image.resize((int(self.PAGE_SIZE[0] * self.DPI_LEVEL), int(self.PAGE_SIZE[1] * self.DPI_LEVEL)))
        return image

    def handle_files(self, files: List[str]) -> str: #处理文件比较流程
        self.logMessage.emit(f"""Processing files:
    {files[0]}
    {files[1]}""")
        toc = []
        current_progress = 0
        with fitz.open(files[0 if self.MAIN_PAGE == "New Document" else 1]) as doc1, fitz.open(
                files[0 if self.MAIN_PAGE == "OLD" else 1]) as doc2:
            size = doc1.load_page(0).rect
            # If page size is auto, self.PAGESIZE will be none
            if self.PAGE_SIZE[0] is None:
                # Assume 72 DPI for original document resolution
                self.PAGE_SIZE = (size.width / 72, size.height / 72)
                self.MERGE_THRESHOLD = int(self.DPI_LEVEL / 120 * self.PAGE_SIZE[0] * self.PAGE_SIZE[1])
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
                stats_filename = path.join(temp_dir, "stats.pdf")
                image_files.append(stats_filename)

                for i in range(total_operations):
                    self.logMessage.emit(f"Processing page {i + 1} of {total_operations}...")
                    self.logMessage.emit(f"Converting main page...")
                    image1 = self.pdf_to_image(i, doc1)
                    self.logMessage.emit(f"Converting secondary page...")
                    image2 = self.pdf_to_image(i, doc2)
                    self.logMessage.emit(f"Marking differences...")
                    markups = self.mark_differences(i, image1, image2)
                    del image1, image2

                    # Save marked images and prepare TOC entries
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

                # Create statistics page
                text = f"Document Comparison Report\n\nTotal Pages: {total_operations}\nFiles Compared:\n    File in Blue_{files[0]}\n    File in Red_{files[1]}\nMain Page: {self.statistics['MAIN_PAGE']}\nTotal Differences: {self.statistics['TOTAL_DIFFERENCES']}\nPages with differences:\n"
                for page_info in self.statistics["PAGES_WITH_DIFFERENCES"]:
                    text += f"    Page {page_info[0] + 1} Changes: {page_info[1]}\n"

                # Create statistics page and handle text overflow
                stats_doc = fitz.open()
                stats_page = stats_doc.new_page()
                text_blocks = text.split('\n')
                y_position = 72
                for line in text_blocks:
                    if y_position > fitz.paper_size('letter')[1] - 72:
                        stats_page = stats_doc.new_page()  # Create a new page if needed
                        y_position = 72  # Reset y position for the new page
                    stats_page.insert_text((72, y_position), line, fontsize=11, fontname="helv")
                    y_position += 12  # Adjust y_position by the line height

                # Save and close the stats document
                stats_filename = path.join(temp_dir, "stats.pdf")
                stats_doc.save(stats_filename)
                stats_doc.close()

                # Builds final PDF from each PDF image page
                self.logMessage.emit("Compiling PDF from output folder...")
                compiled_pdf = fitz.open()
                for img_path in image_files:
                    img = fitz.open(img_path)
                    compiled_pdf.insert_pdf(img, links=False)
                    img.close()

                # Update the table of contents
                compiled_pdf.set_toc(toc)

                # Save Final PDF File
                self.logMessage.emit(f"Saving final PDF...")
                output_path = f"{self.OUTPUT_PATH}{filename.split('.')[0]} Comparison.pdf"
                output_iterator = 0

                # Checks if a version alreaday exists and increments revision if necessary
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


def save_settings(settings: dict) -> None: #保存设置到JSON文件
    settings_path = "settings.json"

    if settings_path:
        with open(settings_path, "w") as f:
            dump(settings, f, indent=4)


def load_settings() -> dict: #从JSON文件加载设置
    settings = None
    settings_path = "settings.json"

    if settings_path and path.exists(settings_path):
        with open(settings_path, "r") as f:
            settings = load(f)
    if not settings:
        settings = _load_default_settings()
    save_settings(settings)
    return settings


def _load_default_settings() -> dict: #加载默认设置
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
