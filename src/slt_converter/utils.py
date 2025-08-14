import os

def valid_qpw(f):
    return f.lower().endswith(".qpw")

def log(msg):
    print(msg)

def batch_files(files):
    valid = [f for f in files if valid_qpw(f)]
    if not valid:
        log("No valid .qpw files found.")
        return []
    log(f"Found {len(valid)} valid .qpw files.")
    return valid
