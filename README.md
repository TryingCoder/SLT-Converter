# AIO File Converter Tool

## Features

- Batch converts `.qpw` files to `.xlsx` using LibreOffice.
- Cleans up leftover `.qpw` files and removes duplicate `.xlsx` files (with parenthesis and numbers).
- Ensures every original file is converted, with final cross-check and reconversion if needed.

## Usage

```sh
python src/converter.py [max_workers]
```
- `max_workers` (optional): Number of parallel conversion threads (default: 4).

## Requirements

- Python 3.x
- LibreOffice installed (`soffice.exe` path found dynamically)
- See `requirements.txt` for dependencies.

## Notes

- Working on expanding functionality to include conversions to different file types, not just limited to .qpw -> .xlsx
- This project was started in an effort solve a problem I encountered while solving an issue for a client
