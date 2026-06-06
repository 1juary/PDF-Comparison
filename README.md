# PDF Toolkit — Compare & Rotate

A unified desktop GUI application for comparing two PDF documents (text diff + visual markup) and rotating PDF pages (WYSIWYG live preview). Both tools are accessible from a single tabbed window — no sub-menus needed.

## Features

### 📄 Compare (Tab 1)
- **Side-by-side file input** — separate drag-and-drop zones for old and new PDF versions
- **Text-based semantic diff** — compares extracted text tokens using sequence matching
- **Visual markup output** — generates a compiled PDF with highlighted differences per page
- **Configurable DPI** — adjustable rendering quality from draft (75 DPI) to professional (1800 DPI)
- **Page size presets** — AUTO, LETTER, ANSI A/B/C/D
- **Output options** — choose which page variants to include (New Copy, Old Copy, Markup, Difference, Overlay)
- **Image formatting** — optional grayscale, black/white, and file size reduction
- **Custom output path** — save results next to source, to a default path, or to a specified directory

### 🔄 Rotate (Tab 2)
- **WYSIWYG rotation** — drag & drop a PDF, see a live 200×200 preview, and rotate with directly visible CW/CCW buttons
- **No intermediate files** — rotation is previewed in memory; no temporary file is written until you explicitly save
- **Per-page navigation** — preview each page before rotating
- **Instant feedback** — rotation angle (0°/90°/180°/270°) is displayed and updated in real time
- **Reset & save** — reset rotation to 0° anytime, or save the rotated PDF to a new file

## Requirements

- Python 3.9+
- PySide6
- PyMuPDF (fitz)
- Pillow (PIL)

```bash
pip install PySide6 PyMuPDF Pillow
```

## Usage

```bash
python main.py
```

### Compare two PDFs
1. Switch to the **📄 Compare** tab
2. Drag an **old version** PDF onto the left panel (or click to browse)
3. Drag a **new version** PDF onto the right panel (or click to browse)
4. Optionally click **Swap** to exchange old/new assignments
5. Select **DPI** and **Page Size** from the dropdowns
6. Click **Compare** to start processing
7. The output PDF is saved next to the main document as `<filename> Comparison.pdf`

### Rotate a PDF
1. Switch to the **🔄 Rotate** tab
2. Drag & drop a PDF onto the drop zone (or click to browse)
3. Use **◀ Prev** / **Next ▶** to navigate pages
4. Click **90° CW** or **90° CCW** to rotate — the preview updates instantly
5. Click **Reset** to revert to 0° rotation
6. Click **Save Rotated PDF** to write the result to disk

## Settings

Click the **⚙ Settings** button to configure:

| Tab | Options |
|-----|---------|
| **Output** | Output path, which page variants to include, scaling, grayscale/BW, file size reduction, main page designation |
| **DPI** | Fine-tune all six DPI presets |
| **Advanced** | Minimum diff token length, text normalization toggle |

## Project Structure

| File | Purpose |
|------|---------|
| `main.py` | **Primary entry point** — unified tabbed GUI combining Compare + Rotate |
| `py_PDF_compare_gui.py` | Comparison engine — text diff, visual markup, settings, and comparison thread |
| `PDF_rotate.py` | Rotation engine — page preview rendering and PDF rotation save logic |
| `PDF_compare_modifiedby_Google_Gemini.py` | **Deprecated** — earlier version with pixel-based comparison (OpenCV), retained for reference only |
| `settings.json` | User settings (auto-generated on first run) |

## Output

The generated comparison PDF includes:
- Per-page visual markup pages (highlighted text differences)
- A structured diff summary with change descriptions
- Optional side-by-side difference views and overlay blends
