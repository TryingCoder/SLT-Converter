# SLT-Converter

A Python-based CLI tool to batch convert `.qpw` files to `.xlsx` using LibreOffice.  
It supports optional backup, duplicate cleanup, and separates failed conversions for easy fault finding.

---

## Features

- Batch convert `.qpw` files to `.xlsx`
- Cleans up duplicate `.qpw` files before conversion
- Separates failed conversions into a `Failed` folder
- CLI-based tool with optional Windows PATH integration
- Cross-platform: Windows, macOS, Linux (requires LibreOffice CLI)

---

## Requirements

- Python 3.10+
- LibreOffice CLI tools (`soffice` executable)
- Python packages: `tqdm`, `pandas`, `openpyxl` (installed automatically via pip)

---

## Installation
```bash
pip install git+https://github.com/TryingCoder/SLT-Converter
```
# Usage
## Normal conversion
```bash
slt --source ./qpw_files --dest ./xlsx_output
```
## Skip duplicates
```bash
slt --source ./qpw_files --dest ./xlsx_output --skip-duplicates
```
## Set concurrency
```bash
slt --source ./qpw_files --dest ./xlsx_output --workers 8
```

## Update to latest version
```bash
slt --update
```
## Show help menu
```bash
slt --help
```
```bash
slt -?
```

## Notes

- Failed conversions are copied to a Failed folder inside the destination directory.

- Temporary working folders are automatically cleaned up.

- Ensure LibreOffice CLI tools are installed and available in your PATH (soffice.exe on Windows, soffice on macOS/Linux).