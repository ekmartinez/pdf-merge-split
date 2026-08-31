# boomPDF

A lightweight Python toolkit for common PDF manipulation tasks — extracting, merging, deleting, inserting, and rotating pages, plus converting PDFs into Word documents.

## Features

- **Select** one or multiple PDF files by path
- **Extract** specific pages (single page, page range, or list of pages) into a new PDF
- **Merge** multiple selected PDFs into a single file
- **Delete** specific pages from a PDF (in place)
- **Insert** a page from one PDF into another at a chosen position
- **Rotate** specific pages by 90°, 180°, or 270°
- **Convert** a PDF (entire document or specific pages) into a Word (`.docx`) file

## Requirements

- Python 3.x
- [`pypdf`](https://pypi.org/project/pypdf/)
- [`pdf2docx`](https://pypi.org/project/pdf2docx/)

Install dependencies:

```bash
pip install pypdf pdf2docx
```

## Usage

```python
from pdf_operations import PDFOperations

pdf_ops = PDFOperations()

# Select the PDF files you'll work with
pdf_ops.select_pdf_files(['ISC_Notes.pdf'])

# Extract pages 1-5 into a new file
pdf_ops.extract_pdf_pages('ISC_Notes.pdf', "1-5", "extracted.pdf")

# Merge all selected PDFs into one
pdf_ops.merge_pdf_files('merged_output.pdf')

# Delete pages 1, 3, and 5
pdf_ops.delete_pdf_pages('merged_output.pdf', [1, 3, 5])

# Insert page 1 of file2.pdf into merged_output.pdf at position 2
pdf_ops.insert_into_pdf('merged_output.pdf', 'file2.pdf', 1, 2)

# Rotate page 1 by 90 degrees
pdf_ops.rotate_pdf_pages('merged_output.pdf', 1, 90)

# Convert the whole PDF to Word
pdf_ops.convert_pdf_to_word('ISC_Notes.pdf', 'file1.docx')

# Convert only specific pages to Word
pdf_ops.convert_pdf_to_word('ISC_Notes.pdf', 'file1_partial.docx', page_specifications="1-3")
```

> **Note:** All page numbers are **1-based** across every method.

## Method reference

| Method | Description |
|---|---|
| `select_pdf_files(paths)` | Registers one or more PDF paths for use by other methods. |
| `get_selected_files()` | Returns the list of currently selected PDF paths. |
| `extract_pdf_pages(pdf_path, page_specifications, output_filename)` | Extracts pages into a new PDF file. |
| `merge_pdf_files(output_filename)` | Merges all selected PDFs into one file. |
| `delete_pdf_pages(pdf_path, page_specifications)` | Deletes pages and **overwrites** the original file. |
| `insert_into_pdf(target_pdf_path, source_pdf_path, page_number, insert_position)` | Inserts a page from one PDF into another. |
| `rotate_pdf_pages(pdf_path, page_specifications, angle)` | Rotates pages by 90/180/270° and **overwrites** the original file. |
| `convert_pdf_to_word(pdf_path, docx_path, page_specifications=None)` | Converts a PDF (or selected pages) to `.docx`. |

## ⚠️ Known limitations

- `delete_pdf_pages`, `rotate_pdf_pages`, and `insert_into_pdf` overwrite the original file in place — **keep a backup** before running them on files you can't lose.
- A file must be registered via `select_pdf_files()` before most operations will run on it.

## Changelog

**v2 — bug fixes and documentation pass**
- Fixed `merge_pdf_files` writing the output file on every loop iteration instead of once at the end.
- Fixed silently-ignored out-of-range page numbers in the single-page-as-string branch of `extract_pdf_pages`, `rotate_pdf_pages`, and `convert_pdf_to_word`.
- Replaced the deprecated `rotate_clockwise()` call with `rotate()` in `rotate_pdf_pages`.
- Fixed several typos in docstrings and error messages.
- Added a class-level docstring.
- Renamed the module from `pdf_procedures.py` to `pdf_operations.py` to match the class it defines.

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.
