import os
import shutil
import re

def validate_file_format(filepath):
    return os.path.isfile(filepath) and filepath.lower().endswith(".qpw")

def create_backup(src_folder, backup_folder):
    if not os.path.exists(backup_folder):
        shutil.copytree(src_folder, backup_folder)

def clean_duplicates(folder, ext=".xlsx"):
    files = [f for f in os.listdir(folder) if f.lower().endswith(ext)]
    seen = set()
    for f in files:
        base = re.sub(r"\(\d+\)", "", os.path.splitext(f)[0]).strip()
        if base.lower() in seen:
            os.remove(os.path.join(folder, f))
        else:
            seen.add(base.lower())

def get_files(folder, ext=".qpw"):
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(ext)
    ]

def retry(func, retries=3, *args, **kwargs):
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt == retries - 1:
                raise
