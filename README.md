# SLT-Converter

A Python-based CLI tool to batch convert `.qpw` files to `.xlsx` --for now functionality is limited to this

---

## Features

- Batch converts `.qpw` files to `.xlsx`
- Cleans up duplicate `.xlsx` files
- Separates failed conversions for easy fault finding
- CLI-based tool
- Optional automatic PATH integration for Windows CLI usage

---

## Requirements

- Python 3.6+
- LibreOffice CLI tools (`soffice.exe`) installed
- Python packages: `tqdm`, `pandas`, `openpyxl` (handled automatically during install)

---

## Installation

```bash
pip install git+https://github.com/TryingCoder/SLT-Converter

## After installation, run:

python -m slt_converter.path_entry

---