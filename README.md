# PDF Comparison Tool

A desktop GUI application for comparing two PDF documents, highlighting text differences and generating comparison reports.

## Features

- **Side-by-side file input** — separate drag-and-drop zones for old and new PDF versions
- **Text-based semantic diff** — compares extracted text tokens using sequence matching
- **Visual markup output** — generates a compiled PDF with highlighted differences per page
- **Configurable DPI** — adjustable rendering quality from draft (75 DPI) to professional (1800 DPI)
- **Page size presets** — AUTO, LETTER, ANSI A/B/C/D
- **Output options** — choose which page variants to include (New Copy, Old Copy, Markup, Difference, Overlay)
- **Image formatting** — optional grayscale, black/white, and file size reduction
- **Custom output path** — save results next to source, to a default path, or to a specified directory

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
python py_PDF_compare_gui.py
```

1. Drag an **old version** PDF onto the left panel (or click to browse)
2. Drag a **new version** PDF onto the right panel (or click to browse)
3. Optionally click **Swap** to exchange old/new assignments
4. Select **DPI** and **Page Size** from the dropdowns
5. Click **Compare** to start processing
6. The output PDF is saved next to the main document as `<filename> Comparison.pdf`

## Settings

Click the **Settings** button in the title bar to configure:

| Tab | Options |
|-----|---------|
| **Output** | Output path, which page variants to include, scaling, grayscale/BW, file size reduction, main page designation |
| **DPI** | Fine-tune all six DPI presets |
| **Advanced** | Minimum diff token length, text normalization toggle |

## Project Structure

| File | Purpose |
|------|---------|
| `py_PDF_compare_gui.py` | **Primary application** — full GUI with text-based PDF comparison |
| `PDF_compare_modifiedby_Google_Gemini.py` | **Deprecated** — earlier version with pixel-based comparison (OpenCV), retained for reference only |
| `settings.json` | User settings (auto-generated on first run) |

## Output

The generated comparison PDF includes:
- Per-page visual markup pages (highlighted text differences)
- A structured diff summary with change descriptions
- Optional side-by-side difference views and overlay blends
