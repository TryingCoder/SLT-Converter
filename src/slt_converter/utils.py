import os
import re

def is_qpw(file):
    return file.lower().endswith(".qpw")

def find_dupes(folder, ext=".qpw"):
    files = [f for f in os.listdir(folder) if f.lower().endswith(ext.lower())]
    base_map = {}
    for f in files:
        match = re.match(r"^(.*?)(?: \((\d+)\))?{}".format(re.escape(ext)), f, re.IGNORECASE)
        if match:
            base = match.group(1).lower()
            num = int(match.group(2)) if match.group(2) else 0
            base_map.setdefault(base, []).append((num, f))
    to_remove = []
    for base, versions in base_map.items():
        versions.sort()
        to_remove.extend(f for _, f in versions[1:])
    return to_remove
