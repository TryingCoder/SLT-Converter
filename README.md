# SLT-Converter v0.1.0 (Beta)

A Python-based CLI tool to batch convert `.qpw` files to `.xlsx` using LibreOffice headless.  
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

See requirements.txt

---

## Installation
```bash
winget install https://github.com/TryingCoder/SLT-Converter/releases/download/v0.1.0/slt_converter.exe
```
# Usage
## Normal conversion
```bash
slt --source ./qpw_files --dest ./xlsx_output
```
## Skip duplicate check
```bash
slt --source ./qpw_files --dest ./xlsx_output --skip-duplicates
```
## Define max workers (Concurrent conversions - Default = 4)
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
slt -h
```

## Notes

- Failed conversions are copied to a Failed folder inside the destination directory.
- Temporary working folders are automatically cleaned up.
- Working on expansion to add more features
- Submit requests to thebrandbackup1@gmail.com or feel free to contribute :D
