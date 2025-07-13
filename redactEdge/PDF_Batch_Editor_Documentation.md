
# 📝 PDF Batch Editor – Engineering Drawing Text & Image Tool

## 📌 Overview

This is a desktop GUI tool built using **Python and Tkinter** that allows users to **batch edit PDF files** (especially engineering drawings). The app supports operations like:

- ✅ Deleting specific text
- 🔁 Replacing text
- 🖼️ Replacing or deleting images
- 🧾 Adding text boxes
- 🧽 Deleting selected table areas

---

## 📦 Project Features

- Multi-file batch processing
- Preview selected PDF pages
- Draw rectangles to specify areas
- Save final files as PDF or JPEG
- Handles watermark removal (Spire.PDF)

---

## 🧾 Dependencies

| Library | Purpose |
|--------|---------|
| `os`, `tempfile` | File path operations and temp file handling |
| `tkinter`, `ttk`, `filedialog`, `messagebox`, `scrolledtext` | GUI creation using Tkinter widgets |
| `PIL (Pillow)` | Display PDF pages as images |
| `spire.pdf` & `spire.pdf.common` | Replace and draw text inside PDFs (for precision rendering) |
| `fitz` (`PyMuPDF`) | Image manipulation, redaction, drawing, and page previews |

---

## ⚙️ Setup Instructions

### ✅ 1. Install Required Python Libraries

You must have Python 3.8+ installed. Then install dependencies:

```bash
pip install Pillow
pip install PyMuPDF
```

### ✅ 2. Install Spire.PDF for Python

- Download from: [https://www.e-iceblue.com/Download/pdf-for-python-free.html](https://www.e-iceblue.com/Download/pdf-for-python-free.html)
- Install the provided `.whl` file:

```bash
pip install Spire.PDF_Free-*.whl
```

> ⚠️ *Note*: Free version of Spire.PDF adds a watermark. This tool removes it by overwriting the top area of pages using `fitz`.

---

## 📂 Folder Structure (if organizing)

```
project/
│
├── pdf_batch_editor.py       # Main Python script (your current file)
├── requirements.txt          # (Optional) for pip install
└── assets/                   # (Optional) for replacement images etc.
```

---

## 💻 How to Run

Run the script with:

```bash
python pdf_batch_editor.py
```

You will see a GUI window allowing you to upload PDFs, choose operations, and process files.

---

## 🛠 Features Summary

| Feature | Method / Class |
|--------|----------------|
| Upload PDFs | `upload_files()` |
| Select output folder | `select_output_dir()` |
| Replace Text | `replace_text_in_pdf()` |
| Delete Text | `delete_text_in_pdf()` |
| Replace Image(s) | `replace_images_in_pdf()` |
| Delete Image(s) | `delete_images_in_pdf()` |
| Add Textbox | `add_textbox_to_pdf()` |
| Delete Table Area | `delete_selected_area()` |
| Preview Pages | `show_pdf_preview()` |
| Save as JPEG | `save_as_jpeg_var`, uses `fitz.get_pixmap()` |

---

## 🧠 Technical Notes

- **fitz (PyMuPDF)** is used for page rendering, area deletion, image redaction, and JPEG conversion.
- **Spire.PDF** is required for text replacement and text box drawing.
- Drawing areas for text/table deletion is done via mouse interaction on `Canvas`.
- All intermediate steps are saved as temporary files and replaced in final output.
- `VerticalScrolledFrame` is a custom scrollable container for large GUI layout.

---

## 📌 Requirements Summary

| Requirement | Version |
|-------------|---------|
| Python | 3.8+ |
| PyMuPDF (`fitz`) | Latest |
| Pillow (`PIL`) | Latest |
| Spire.PDF Free | Required (download manually) |

---

## ✅ To-Do / Enhancements (Optional)

- Add undo/redo feature for previews
- Support custom font and color for text box
- Export table areas as CSV
- Add support for drag-and-drop file input
