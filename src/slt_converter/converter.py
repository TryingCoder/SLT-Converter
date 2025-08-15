import os
import re
import shutil
import subprocess
import sys
import tempfile
import importlib.util
import site
from concurrent.futures import ThreadPoolExecutor, as_completed
from slt_converter.utils import is_qpw, find_dupes

# -----------------------------
# Ensure tqdm installed
# -----------------------------
def ensure_tqdm():
    if importlib.util.find_spec("tqdm") is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--no-warn-script-location", "tqdm"])
    user_site = site.getusersitepackages()
    if user_site not in sys.path and os.path.exists(user_site):
        sys.path.append(user_site)
    global tqdm
    from tqdm import tqdm

ensure_tqdm()

# -----------------------------
# LibreOffice detection
# -----------------------------
def find_soffice():
    paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    raise FileNotFoundError("LibreOffice CLI not found. Install LibreOffice CLI tools.")

SOFFICE = find_soffice()

# -----------------------------
# Utility functions
# -----------------------------
def mkdir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def copy_files(src, dst, ext=".qpw", desc="Copying"):
    files = [f for f in os.listdir(src) if f.lower().endswith(ext.lower())]
    mkdir(dst)
    from tqdm import tqdm
    with tqdm(total=len(files), desc=desc, unit="file") as pbar:
        for f in files:
            shutil.copy2(os.path.join(src, f), os.path.join(dst, f))
            pbar.update(1)

def cleanup_dupes(folder, ext=".qpw", prompt=True):
    dupes = find_dupes(folder, ext)
    if dupes and prompt:
        print("\nDuplicates found:")
        for f in dupes:
            print(" -", f)
        if input("Remove duplicates? (Y/N): ").strip().lower() in ["y","yes"]:
            from tqdm import tqdm
            with tqdm(total=len(dupes), desc="Removing Duplicates") as pbar:
                for f in dupes:
                    os.remove(os.path.join(folder, f))
                    pbar.update(1)
            return len(dupes)
    return 0

def convert_file(src_file, out_dir):
    out_file = os.path.join(out_dir, os.path.basename(src_file).replace(".qpw",".xlsx"))
    if os.path.exists(out_file):
        return True
    try:
        cmd = [SOFFICE, "--headless", "--convert-to", "xlsx", src_file, "--outdir", out_dir]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return os.path.exists(out_file)
    except Exception:
        return False

def convert_all(src_folder, out_folder, failed_folder, max_workers=4):
    mkdir(failed_folder)
    all_files = [f for f in os.listdir(src_folder) if is_qpw(f)]
    pending = set(all_files)
    succeeded = set()
    last_count = len(pending)
    from tqdm import tqdm
    with tqdm(total=len(all_files), desc="Converting", unit="file") as pbar:
        attempt = 1
        while pending:
            current = list(pending)
            failed_this_round = set()
            with ThreadPoolExecutor(max_workers=max_workers) as exe:
                futures = {exe.submit(convert_file, os.path.join(src_folder, f), out_folder): f for f in current}
                for fut in as_completed(futures):
                    f = futures[fut]
                    if fut.result():
                        if f not in succeeded:
                            pbar.update(1)
                            succeeded.add(f)
                            pending.discard(f)
                        fp = os.path.join(failed_folder, f)
                        if os.path.exists(fp):
                            os.remove(fp)
                    else:
                        failed_this_round.add(f)
                        dst = os.path.join(failed_folder, f)
                        if not os.path.exists(dst):
                            shutil.copy2(os.path.join(src_folder, f), dst)
            if not failed_this_round or len(failed_this_round) == last_count:
                break
            pending = failed_this_round
            last_count = len(pending)
            attempt += 1
    return succeeded, pending

# -----------------------------
# Main CLI
# -----------------------------
def main():
    src = input("Enter SOURCE folder path: ").strip()
    while not os.path.isdir(src):
        src = input("Invalid path. Enter SOURCE folder path: ").strip()
    dst = input("Enter DESTINATION folder path: ").strip()
    mkdir(dst)
    failed = os.path.join(dst, "Failed")
    mkdir(failed)
    max_workers = int(sys.argv[1]) if len(sys.argv) > 1 else 4
    temp_dir = tempfile.mkdtemp(prefix="qpw_work_")
    print(f"Temp working directory: {temp_dir}")
    copy_files(src, temp_dir, desc="Populating")
    dup_removed = cleanup_dupes(temp_dir)
    print(f"Converting {len([f for f in os.listdir(temp_dir) if is_qpw(f)])} files...")
    succeeded, failed_files = convert_all(temp_dir, dst, failed, max_workers)
    shutil.rmtree(temp_dir)
    print(f"Temp folder removed: {temp_dir}")
    print("\n===== Summary =====")
    print(f"✅ Total files: {len([f for f in os.listdir(src) if is_qpw(f)])}")
    print(f"✅ Converted: {len(succeeded)}")
    print(f"❌ Failed: {len(failed_files)}")
    print(f"✅ Duplicates removed: {dup_removed}")
    if not failed_files and os.path.exists(failed):
        os.rmdir(failed)
    print("===================")
    print("All done!")

