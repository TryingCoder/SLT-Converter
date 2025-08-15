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
pip install git+https://github.com/TryingCoder/SLT-Converter.git

---

## Windows Setup

After installation:

The installer checks if the Python Scripts folder (where CLI commands are installed) is in your PATH.

If missing, you'll see a prompt:
Add slt-converter to PATH? (y/n)