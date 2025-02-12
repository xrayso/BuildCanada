# PDF Table Extractor

This script extracts tables from a PDF and saves each table to a separate sheet in an Excel file.

## Features
- Detects tables using `Camelot`
- Extracts and saves tables into an Excel file
- Supports multiple tables per page

## Installation
```bash
pip install fitz PyMuPDF camelot-py[cv] pandas openpyxl
