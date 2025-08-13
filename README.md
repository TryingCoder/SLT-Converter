# QPW to XLSX Converter

## Features

- Batch converts `.qpw` files to `.xlsx` using LibreOffice.
- Automatically retries failed conversions until all possible files are processed.
- Cleans up leftover `.qpw` files and removes duplicate `.xlsx` files (with parenthesis and numbers).
- Ensures every original file is converted, with final cross-check and reconversion if needed.

## Usage

```sh
python src/converter.py [max_workers]
```
- `max_workers` (optional): Number of parallel conversion threads (default: 4).

## Cleanup Logic

- **Leftover .qpw files** in `Merz_Converted` are deleted.
- **Duplicate .xlsx files** (e.g., `filename (1).xlsx`, `filename (2).xlsx`) are removed if the original exists.
- **Failed conversions** are retried until no `.qpw` files remain in `Merz_Failed`.

## Requirements

- Python 3.x
- LibreOffice installed (`soffice.exe` path must be correct)
- See `requirements.txt` for dependencies.

## Folder Structure

- `Merz_Original`: Source `.qpw` files
- `Merz_TBC`: Temporary working folder
- `Merz_Failed`: Files that failed conversion
- `Merz_Converted`: Successfully converted `.xlsx` files

## Notes

- The script is optimized for speed and reliability.
- All cleanup and retry logic is automatic.