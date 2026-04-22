# PDF-Comparison

The script to compare the difference between PDFs based on Python.

## Key Features Implemented

### Content Extraction
- Extracts text blocks and potential drawing regions from PDFs using PyMuPDF and OpenCV.

### Text Similarity Matching
- Uses difflib to compare text blocks across pages, allowing for minor differences (configurable threshold).

### Drawing Similarity Matching
- Uses ORB feature detection to match image regions based on shape and local features.

### Cross-Page Matching
- Compares content across all pages, not just page-by-page.

## Output Options

### Report Mode
- Generates a detailed text report of matches in the statistics page.

### Markup Mode
- Framework in place (pixel-based markup still works; similarity highlights planned for future).
