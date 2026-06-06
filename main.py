import os
import sys
import tempfile
import shutil
import fitz
from PIL import Image
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QImage
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QStyleFactory,
    QVBoxLayout,
    QWidget,
    QMessageBox,
)

# 导入原始脚本中的组件
import py_PDF_compare_gui
from py_PDF_compare_gui import (
    MainWindow,
    DragDropLabel,
    stylesheet,
)

class ExtendedMainWindow(MainWindow):
    def __init__(self):
        super().__init__()
        # 初始化我们新增的预览与旋转小部件
        self.init_extended_widgets()
        # 重构整体界面布局
        self.rebuild_layout()
        # 刷新界面状态
        self.update_ui_state()

    def init_extended_widgets(self):
        """初始化在主界面展示的预览区与一级菜单旋转按钮"""
        # --- 旧文档区域组件 ---
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

        old_btn_style = self._btn_style("#FF6B6B", "#FFF5F5", "#FFE0E0")
        self.rot_ccw_old_btn = QPushButton("↺ 90° CCW")
        self.rot_ccw_old_btn.setStyleSheet(old_btn_style)
        self.rot_ccw_old_btn.clicked.connect(lambda: self.rotate_file_in_place("old", -90))

        self.rot_cw_old_btn = QPushButton("90° CW ↻")
        self.rot_cw_old_btn.setStyleSheet(old_btn_style)
        self.rot_cw_old_btn.clicked.connect(lambda: self.rotate_file_in_place("old", 90))

        # --- 新文档区域组件 ---
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

        new_btn_style = self._btn_style("#2196F3", "#F0F8FF", "#E3F2FD")
        self.rot_ccw_new_btn = QPushButton("↺ 90° CCW")
        self.rot_ccw_new_btn.setStyleSheet(new_btn_style)
        self.rot_ccw_new_btn.clicked.connect(lambda: self.rotate_file_in_place("new", -90))

        self.rot_cw_new_btn = QPushButton("90° CW ↻")
        self.rot_cw_new_btn.setStyleSheet(new_btn_style)
        self.rot_cw_new_btn.clicked.connect(lambda: self.rotate_file_in_place("new", 90))

    def rebuild_layout(self):
        """丢弃原有 MainWindow 的旧垂直布局，重新编排一个整洁的紧凑型双栏布局"""
        # 调整主窗口尺寸以完美适应双栏直观预览
        self.resize(550, 600)

        # 核心中央小部件
        new_central = QWidget()
        new_layout = QVBoxLayout(new_central)
        new_layout.setContentsMargins(16, 16, 16, 16)
        new_layout.setSpacing(12)

        # 1. 顶部左右双栏：左边 Old / 右边 New
        cols_layout = QHBoxLayout()
        cols_layout.setSpacing(16)

        # --- 左侧一栏 (Old) ---
        old_col = QVBoxLayout()
        old_col.setSpacing(8)
        self.drop_label_old.setMaximumHeight(50)  # 适当压缩拖拽栏，为预览预留空间 [1]
        self.drop_label_old.setMinimumHeight(45)
        old_col.addWidget(self.drop_label_old)
        
        # 预览容器居中
        preview_old_container = QHBoxLayout()
        preview_old_container.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_old_container.addWidget(self.preview_old)
        old_col.addLayout(preview_old_container)

        # 旋转控制横排
        old_rot_layout = QHBoxLayout()
        old_rot_layout.setSpacing(6)
        old_rot_layout.addWidget(self.rot_ccw_old_btn)
        old_rot_layout.addWidget(self.rot_cw_old_btn)
        old_col.addLayout(old_rot_layout)

        # --- 右侧一栏 (New) ---
        new_col = QVBoxLayout()
        new_col.setSpacing(8)
        self.drop_label_new.setMaximumHeight(50)
        self.drop_label_new.setMinimumHeight(45)
        new_col.addWidget(self.drop_label_new)

        # 预览容器居中
        preview_new_container = QHBoxLayout()
        preview_new_container.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_new_container.addWidget(self.preview_new)
        new_col.addLayout(preview_new_container)

        # 旋转控制横排
        new_rot_layout = QHBoxLayout()
        new_rot_layout.setSpacing(6)
        new_rot_layout.addWidget(self.rot_ccw_new_btn)
        new_rot_layout.addWidget(self.rot_cw_new_btn)
        new_col.addLayout(new_rot_layout)

        cols_layout.addLayout(old_col)
        cols_layout.addLayout(new_col)
        new_layout.addLayout(cols_layout)

        # 2. 中部功能按键：Swap 按钮 与 Compare 按钮
        new_layout.addWidget(self.swap_button)
        new_layout.addWidget(self.compare_button)

        # 3. 底部配置行（将原版一整行的 DPI 和 Page Size 并排，节省垂直空间）
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

        self.setCentralWidget(new_central)

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

    def refresh_preview(self, role):
        """即时加载并渲染 PDF 的第一页"""
        label = self.drop_label_new if role == "new" else self.drop_label_old
        preview_widget = self.preview_new if role == "new" else self.preview_old
        file_path = label.file_path

        if not file_path or not os.path.exists(file_path):
            preview_widget.clear()
            text = "Drop NEW PDF here\nor click to browse" if role == "new" else "Drop OLD PDF here\nor click to browse"
            preview_widget.setText(text)
            return

        try:
            # 打开、渲染、随后立即关闭文档（防止文件占用锁定，方便本地写入和旋转）
            doc = fitz.open(file_path)
            page = doc.load_page(0)
            
            rect = page.rect
            scale = 210 / max(rect.width, rect.height, 1)  # 缩放到预览框内
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
            
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            doc.close()

            data = img.tobytes("raw", "RGB")
            qt_img = QImage(data, img.width, img.height, img.width * 3, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qt_img)
            preview_widget.setPixmap(pixmap)
        except Exception as e:
            preview_widget.setText(f"Preview Error:\n{e}")

    def rotate_file_in_place(self, role, angle_diff):
        """直接旋转本地的真实文件"""
        label = self.drop_label_new if role == "new" else self.drop_label_old
        file_path = label.file_path
        if not file_path or not os.path.exists(file_path):
            return

        try:
            # 1. 打开 PDF
            doc = fitz.open(file_path)
            # 2. 遍历并顺时针/逆时针旋转每一页 [1]
            for page in doc:
                page.set_rotation((page.rotation + angle_diff) % 360)
            
            # 3. 为防止覆写时发生进程占用，将修改安全写入临时文件后，再替换原真实文件 [1]
            temp_fd, temp_path = tempfile.mkstemp(suffix=".pdf")
            os.close(temp_fd)
            doc.save(temp_path)
            doc.close()

            shutil.move(temp_path, file_path)

            # 4. 刷新当前的预览图 [1]
            self.refresh_preview(role)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to modify and save PDF locally:\n{e}")

    def update_ui_state(self):
        """动态控制按钮的可点击状态与预览刷新"""
        has_old = bool(self.drop_label_old.file_path and os.path.exists(self.drop_label_old.file_path))
        has_new = bool(self.drop_label_new.file_path and os.path.exists(self.drop_label_new.file_path))

        # 启用/禁用旋转按钮
        self.rot_ccw_old_btn.setEnabled(has_old)
        self.rot_cw_old_btn.setEnabled(has_old)
        self.rot_ccw_new_btn.setEnabled(has_new)
        self.rot_cw_new_btn.setEnabled(has_new)

        # 刷新两侧预览
        self.refresh_preview("old")
        self.refresh_preview("new")

    def swap_files(self):
        """交换文件，交换后自动更新状态"""
        super().swap_files()
        self.update_ui_state()


# ---------------------------------------------------------------------------
# 拦截 DragDropLabel 注入状态刷新
# ---------------------------------------------------------------------------
_orig_set_file = DragDropLabel.set_file

def patched_set_file(self, file_path):
    """当用户拖入、切换或选择文件时，自动更新一级页面的预览和旋转按钮状态"""
    _orig_set_file(self, file_path)
    if hasattr(self, "_parent") and self._parent:
        if hasattr(self._parent, "update_ui_state"):
            self._parent.update_ui_state()

DragDropLabel.set_file = patched_set_file


# ---------------------------------------------------------------------------
# 启动入口
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication([])
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setStyleSheet(stylesheet)

    window = ExtendedMainWindow()
    
    # 动态将当前实例注入到 py_PDF_compare_gui 命名空间下，防 DPISettings 报错
    py_PDF_compare_gui.window = window

    window.show()
    app.exec()