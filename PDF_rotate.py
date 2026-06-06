"""
PDF Page Rotation Tool

Drag-and-drop a PDF, preview pages in a 200×200 window,
rotate clockwise or counter-clockwise, then save.
"""

import sys
from os import path

import fitz
from PIL import Image, ImageQt
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QPixmap, QImage
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QStyleFactory,
    QVBoxLayout,
    QWidget,
)


# ---------------------------------------------------------------------------
# light wrappers over PyMuPDF + Pillow
# ---------------------------------------------------------------------------

def render_page_preview(pdf_path: str, page_index: int, rotation: int,
                        preview_size: int = 200) -> Image.Image:
    """Render one page fitted inside a ``preview_size × preview_size`` box
    **preserving the original aspect ratio**, rotated by *rotation* degrees."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_index)
    
    # cardinal rotation swaps width/height, so determine the "rotated" dimensions
    rot = rotation % 360
    pw, ph = page.rect.width, page.rect.height
    if rot in (90, 270):
        pw, ph = ph, pw

    # scale so the longer side fits inside preview_size
    scale = preview_size / max(pw, ph, 1)
    
    # [FIXED] Pass scale for both X and Y axes. 
    # Passing a single argument creates a rotation matrix, not a scale matrix!
    mat = fitz.Matrix(scale, scale)

    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    doc.close()

    rotated = img.rotate(-rot, expand=True, fillcolor=(255, 255, 255))

    # centre inside a preview_size × preview_size canvas (letterboxed)
    canvas = Image.new("RGB", (preview_size, preview_size), (255, 255, 255))
    px = (preview_size - rotated.width) // 2
    py = (preview_size - rotated.height) // 2
    canvas.paste(rotated, (px, py))
    return canvas


def rotate_pdf(source_path: str, out_path: str, rotation: int):
    """Apply *rotation* (in degrees, multiple of 90) to every page and save."""
    doc = fitz.open(source_path)
    for page in doc:
        page.set_rotation((page.rotation + rotation) % 360)
    doc.save(out_path)
    doc.close()


# ---------------------------------------------------------------------------
# Custom title bar  (mirrors the main app's look)
# ---------------------------------------------------------------------------

class TitleBar(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self._parent = parent
        self.setFixedHeight(40)

        self._drag_start: QPoint | None = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 0, 10, 0)

        self.title_label = QLabel("PDF Rotation Tool")
        self.title_label.setStyleSheet(
            "color: white; font: bold 14px 'Segoe UI', Arial, sans-serif;"
        )

        spacer = QSpacerItem(20, 20, QSizePolicy.Policy.Expanding,
                             QSizePolicy.Policy.Minimum)

        self.min_btn = QPushButton("−", self)
        self.min_btn.setObjectName("MinimizeButton")
        self.min_btn.setFixedSize(28, 28)
        self.min_btn.clicked.connect(self._parent.showMinimized)

        self.close_btn = QPushButton("✕", self)
        self.close_btn.setObjectName("CloseButton")
        self.close_btn.setFixedSize(28, 28)
        self.close_btn.clicked.connect(self._parent.close)

        layout.addItem(spacer)
        layout.addWidget(self.title_label)
        layout.addItem(spacer)
        layout.addWidget(self.min_btn)
        layout.addWidget(self.close_btn)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start = event.globalPosition().toPoint() - self._parent.pos()
        event.accept()

    def mouseMoveEvent(self, event):
        if self._drag_start is not None and event.buttons() & Qt.LeftButton:
            self._parent.move(event.globalPosition().toPoint() - self._drag_start)
        event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_start = None
        event.accept()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class RotateWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Rotation Tool")
        self.setGeometry(200, 200, 420, 480)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)

        # ---- state ----
        self._pdf_path: str | None = None
        self._page_count: int = 0
        self._current_page: int = 0
        self._rotation: int = 0  # cumulative rotation in degrees (0 | 90 | 180 | 270)

        # ---- title bar ----
        self.title_bar = TitleBar(self)
        self.title_bar.setObjectName("TitleBar")
        self.setMenuWidget(self.title_bar)

        # ---- central widget ----
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        # ----- drop zone / file picker -----
        self.drop_btn = QPushButton("Drop PDF here\nor click to browse")
        self.drop_btn.setAcceptDrops(True)
        self.drop_btn.clicked.connect(self._browse)
        self.drop_btn.dragEnterEvent = self._drag_enter
        self.drop_btn.dropEvent = self._drop
        self.drop_btn.setStyleSheet(self._drop_style("#2196F3", "#F0F8FF", "#E3F2FD"))
        self.drop_btn.setMinimumHeight(80)
        root.addWidget(self.drop_btn)

        # ----- page navigation -----
        nav_row = QHBoxLayout()
        self.prev_btn = QPushButton("◀  Prev")
        self.prev_btn.clicked.connect(self._prev_page)
        self.prev_btn.setEnabled(False)
        nav_row.addWidget(self.prev_btn)

        self.page_label = QLabel("Page 0 / 0")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_label.setStyleSheet(
            "font: 13px 'Segoe UI', Arial, sans-serif; color: #1A1A2E;"
        )
        nav_row.addWidget(self.page_label)

        self.next_btn = QPushButton("Next  ▶")
        self.next_btn.clicked.connect(self._next_page)
        self.next_btn.setEnabled(False)
        nav_row.addWidget(self.next_btn)
        root.addLayout(nav_row)

        # ----- preview (200 × 200 px) -----
        self.preview_label = QLabel()
        self.preview_label.setFixedSize(200, 200)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setStyleSheet(
            "background-color: #FFFFFF;"
            "border: 1px solid #E0E6ED;"
            "border-radius: 8px;"
        )
        preview_wrapper = QHBoxLayout()
        preview_wrapper.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_wrapper.addWidget(self.preview_label)
        root.addLayout(preview_wrapper)

        # ----- rotation info -----
        self.rot_label = QLabel("Rotation: 0°")
        self.rot_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.rot_label.setStyleSheet(
            "font: bold 16px 'Segoe UI', Arial, sans-serif; color: #1A1A2E;"
        )
        root.addWidget(self.rot_label)

        # ----- rotate buttons -----
        rot_row = QHBoxLayout()

        self.ccw_btn = QPushButton("↺  90° CCW")
        self.ccw_btn.clicked.connect(self._rotate_ccw)
        self.ccw_btn.setEnabled(False)
        self.ccw_btn.setStyleSheet(self._action_btn_style("#FF6B6B", "#FFF5F5", "#FFE0E0"))
        rot_row.addWidget(self.ccw_btn)

        self.cw_btn = QPushButton("90° CW  ↻")
        self.cw_btn.clicked.connect(self._rotate_cw)
        self.cw_btn.setEnabled(False)
        self.cw_btn.setStyleSheet(self._action_btn_style("#2196F3", "#F0F8FF", "#E3F2FD"))
        rot_row.addWidget(self.cw_btn)
        root.addLayout(rot_row)

        # ----- reset & save -----
        action_row = QHBoxLayout()
        self.reset_btn = QPushButton("Reset")
        self.reset_btn.clicked.connect(self._reset_rotation)
        self.reset_btn.setEnabled(False)
        action_row.addWidget(self.reset_btn)

        self.save_btn = QPushButton("Save Rotated PDF")
        self.save_btn.clicked.connect(self._save)
        self.save_btn.setEnabled(False)
        self.save_btn.setStyleSheet(self._action_btn_style(
            "#4CAF50", "#F1F8E9", "#E8F5E9"
        ))
        action_row.addWidget(self.save_btn)
        root.addLayout(action_row)

        # ---- global stylesheet ----
        self._apply_stylesheet()

    # ------------------------------------------------------------------ #
    #  styles                                                            #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _drop_style(color: str, bg: str, hover_bg: str) -> str:
        return f"""
        QPushButton {{
            color: {color};
            background-color: {bg};
            border-radius: 12px;
            border: 2px dashed {color};
            font: bold 13px "Segoe UI", Arial, sans-serif;
            padding: 8px;
        }}
        QPushButton:hover {{
            background-color: {hover_bg};
            border: 2px solid {color};
        }}
        """

    @staticmethod
    def _action_btn_style(color: str, bg: str, hover_bg: str) -> str:
        return f"""
        QPushButton {{
            color: {color};
            background-color: {bg};
            border: 1px solid {color};
            border-radius: 8px;
            padding: 10px;
            font: bold 14px "Segoe UI", Arial, sans-serif;
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

    def _apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F0F4F8;
            }
            #TitleBar {
                background-color: #2196F3;
            }
            #MinimizeButton, #CloseButton {
                background-color: transparent;
                color: white;
                font-weight: bold;
                border-radius: 4px;
            }
            #MinimizeButton:hover {
                background-color: rgba(255,255,255,0.2);
            }
            #CloseButton:hover {
                background-color: #FF5252;
            }
        """)

    # ------------------------------------------------------------------ #
    #  I/O & drop handling                                               #
    # ------------------------------------------------------------------ #

    def _load_pdf(self, file_path: str):
        """Open a PDF and reset state for the new file."""
        self._pdf_path = file_path
        doc = fitz.open(file_path)
        self._page_count = doc.page_count
        doc.close()

        self._current_page = 0
        self._rotation = 0
        self._update_ui_for_pdf()
        self._redraw_preview()
        self._enable_controls(True)

    def _browse(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open PDF", "", "PDF Files (*.pdf)"
        )
        if file_path:
            self._load_pdf(file_path)

    def _drag_enter(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _drop(self, event):
        urls = event.mimeData().urls()
        if urls:
            p = urls[0].toLocalFile()
            if p.lower().endswith(".pdf"):
                self._load_pdf(p)

    # ------------------------------------------------------------------ #
    #  navigation                                                        #
    # ------------------------------------------------------------------ #

    def _prev_page(self):
        if self._page_count == 0:
            return
        self._current_page = (self._current_page - 1) % self._page_count
        self._redraw_preview()
        self._update_page_label()

    def _next_page(self):
        if self._page_count == 0:
            return
        self._current_page = (self._current_page + 1) % self._page_count
        self._redraw_preview()
        self._update_page_label()

    # ------------------------------------------------------------------ #
    #  rotation                                                          #
    # ------------------------------------------------------------------ #

    def _rotate_cw(self):
        self._rotation = (self._rotation + 90) % 360
        self._redraw_preview()
        self._update_rot_label()

    def _rotate_ccw(self):
        self._rotation = (self._rotation - 90) % 360
        self._redraw_preview()
        self._update_rot_label()

    def _reset_rotation(self):
        self._rotation = 0
        self._redraw_preview()
        self._update_rot_label()

    # ------------------------------------------------------------------ #
    #  save                                                              #
    # ------------------------------------------------------------------ #

    def _save(self):
        if not self._pdf_path:
            return
        base, _ = path.splitext(self._pdf_path)
        default_name = f"{base}_rotated_{self._rotation}.pdf"
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save Rotated PDF", default_name, "PDF Files (*.pdf)"
        )
        if out_path:
            rotate_pdf(self._pdf_path, out_path, self._rotation)
            # update the drop-zone label to show the new file name
            self.drop_btn.setText(
                f"Saved: {path.basename(out_path)}"
            )

    # ------------------------------------------------------------------ #
    #  preview rendering                                                 #
    # ------------------------------------------------------------------ #

    def _redraw_preview(self):
        if not self._pdf_path:
            self.preview_label.clear()
            return
        try:
            pil_img = render_page_preview(
                self._pdf_path, self._current_page, self._rotation, 200
            )
            data = pil_img.tobytes("raw", "RGB")
            qt_img = QImage(data, pil_img.width, pil_img.height,
                            pil_img.width * 3, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qt_img)
            self.preview_label.setPixmap(pixmap)
        except Exception as exc:
            self.preview_label.setText(f"Preview error:\n{exc}")

    # ------------------------------------------------------------------ #
    #  UI helpers                                                        #
    # ------------------------------------------------------------------ #

    def _update_ui_for_pdf(self):
        self.drop_btn.setText(path.basename(self._pdf_path))
        self.drop_btn.setStyleSheet(
            self._drop_style("#4CAF50", "#F1F8E9", "#E8F5E9")
        )
        self._update_page_label()
        self._update_rot_label()

    def _update_page_label(self):
        self.page_label.setText(
            f"Page {self._current_page + 1} / {self._page_count}"
        )

    def _update_rot_label(self):
        self.rot_label.setText(f"Rotation: {self._rotation}°")

    def _enable_controls(self, enabled: bool):
        self.prev_btn.setEnabled(enabled)
        self.next_btn.setEnabled(enabled)
        self.ccw_btn.setEnabled(enabled)
        self.cw_btn.setEnabled(enabled)
        self.reset_btn.setEnabled(enabled)
        self.save_btn.setEnabled(enabled)


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = QApplication([])
    app.setStyle(QStyleFactory.create("Fusion"))

    window = RotateWindow()
    window.show()
    app.exec()