import os
import sys
import shutil
import subprocess
import tempfile
import importlib.util
import site
from concurrent.futures import ThreadPoolExecutor, as_completed
from .utils import validate_file_format

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
# Find soffice.exe
# -----------------------------
def find_soffice():
    paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    try:
        if subprocess.run(["soffice", "--version"], capture_output=True).returncode == 0:
            return "soffice"
    except FileNotFoundError:
        pass
    return None

SOFFICE = find_soffice()
if not SOFFICE:
    print("LibreOffice CLI not found. Install soffice.exe (headless CLI).")
    sys.exit(1)

# -----------------------------
# Utilities
# -----------------------------
def mkdir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def copy(src, dst, ext=".qpw", desc="Copying"):
    files = [f for f in os.listdir(src) if f.lower().endswith(ext.lower())]
    mkdir(dst)
    with tqdm(total=len(files), desc=desc, unit="file") as pbar:
        for f in files:
            shutil.copy2(os.path.join(src, f), os.path.join(dst, f))
            pbar.update(1)

def duplicates(folder, ext=".qpw"):
    import re
    files = [f for f in os.listdir(folder) if f.lower().endswith(ext.lower())]
    base_map = {}
    for f in files:
        m = re.match(r"^(.*?)(?: \((\d+)\))?{}".format(re.escape(ext)), f, re.IGNORECASE)
        if m:
            base = m.group(1).lower()
            num = int(m.group(2)) if m.group(2) else 0
            base_map.setdefault(base, []).append((num, f))
    to_remove = []
    for v in base_map.values():
        v.sort()
        to_remove.extend([fname for _, fname in v[1:]])
    return to_remove

def cleanup_dupes(folder, ext=".qpw", prompt=True):
    dups = duplicates(folder, ext)
    if not dups:
        if prompt:
            print("No duplicates found.")
        return 0
    if prompt:
        print("\nDuplicates:\n" + "\n".join(f" - {f}" for f in dups))
        if input("Remove duplicates? (Y/N): ").strip().lower() in ["y", "yes"]:
            with tqdm(total=len(dups), desc="Removing Duplicates", unit="file") as pbar:
                for f in dups:
                    os.remove(os.path.join(folder, f))
                    pbar.update(1)
            return len(dups)
        return 0
    return 0

def convert_file(src, dst):
    out = os.path.join(dst, os.path.basename(src).replace(".qpw", ".xlsx"))
    if os.path.exists(out):
        return True
    try:
        subprocess.run([SOFFICE, "--headless", "--convert-to", "xlsx", src, "--outdir", dst],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return os.path.exists(out)
    except Exception:
        return False

def convert_all(src, dst, fail, workers=4):
    mkdir(fail)
    files = [f for f in os.listdir(src) if f.lower().endswith(".qpw")]
    pending = set(files)
    done = set()
    last_count = len(pending)

    with tqdm(total=len(files), desc="Converting", unit="file") as pbar:
        attempt = 1
        while pending:
            fail_this_round = set()
            with ThreadPoolExecutor(max_workers=workers) as exec:
                futures = {exec.submit(convert_file, os.path.join(src, f), dst): f for f in pending}
                for fut in as_completed(futures):
                    f = futures[fut]
                    if fut.result():
                        if f not in done:
                            pbar.update(1)
                            done.add(f)
                        pending.discard(f)
                        fp = os.path.join(fail, f)
                        if os.path.exists(fp):
                            os.remove(fp)
                    else:
                        fail_this_round.add(f)
                        fp = os.path.join(fail, f)
                        if not os.path.exists(fp):
                            shutil.copy2(os.path.join(src, f), fp)

            if not fail_this_round or len(fail_this_round) == last_count:
                if fail_this_round:
                    print(f"No progress after attempt {attempt}, {len(fail_this_round)} files failed.")
                break
            pending = fail_this_round
            last_count = len(pending)
            attempt += 1
    return done, pending

def backup(src, dst, ext=".qpw"):
    mkdir(dst)
    existing = [f for f in os.listdir(src) if f.lower().endswith(ext.lower()) and os.path.exists(os.path.join(dst, f))]
    action = "replace"
    if existing:
        print("\nFiles exist in backup:\n" + "\n".join(f" - {f}" for f in existing))
        while True:
            c = input("Skip or Replace? [S/R]: ").strip().lower()
            if c in ["s", "skip"]:
                action = "skip"
                break
            if c in ["r", "replace"]:
                action = "replace"
                break
    with tqdm(total=len([f for f in os.listdir(src) if f.lower().endswith(ext.lower())]), desc="Backup", unit="file") as pbar:
        for f in os.listdir(src):
            if not f.lower().endswith(ext.lower()):
                continue
            src_f = os.path.join(src, f)
            dst_f = os.path.join(dst, f)
            if os.path.exists(dst_f) and action == "replace":
                shutil.copy2(src_f, dst_f)
            elif not os.path.exists(dst_f):
                shutil.copy2(src_f, dst_f)
            pbar.update(1)

# -----------------------------
# Main
# -----------------------------
def main():
    if "--uninstall" in sys.argv:
        import pkg_resources
        dist = pkg_resources.get_distribution("slt-converter")
        print(f"Uninstalling from {dist.location}")
        subprocess.run([sys.executable, "-m", "pip", "uninstall", "-y", "slt-converter"])
        sys.exit(0)

    src_dir = input("SOURCE folder: ").strip()
    while not os.path.isdir(src_dir):
        src_dir = input("Invalid. SOURCE folder: ").strip()

    dst_dir = input("DESTINATION folder: ").strip()
    mkdir(dst_dir)
    fail_dir = os.path.join(dst_dir, "Failed")
    mkdir(fail_dir)
    workers = int(sys.argv[1]) if len(sys.argv) > 1 else 4

    if input("Create backup? (Y/N): ").strip().lower() in ["y", "yes"]:
        bkp_dir = input("BACKUP folder: ").strip()
        backup(src_dir, bkp_dir)

    tmp_dir = tempfile.mkdtemp(prefix="qpw_work_")
    copy(src_dir, tmp_dir, desc="Populating")
    dup_count = cleanup_dupes(tmp_dir, prompt=True)
    print(f"\nConverting {len([f for f in os.listdir(tmp_dir) if f.lower().endswith('.qpw')])} files...")
    done, failed = convert_all(tmp_dir, dst_dir, fail_dir, workers)
    shutil.rmtree(tmp_dir)

    total = len([f for f in os.listdir(src_dir) if f.lower().endswith(".qpw")])
    print(f"\n===== Summary =====\nTotal: {total}\nConverted: {len(done)}\nFailed: {len(failed)}\nDuplicates Removed: {dup_count}")
    if not failed and os.path.exists(fail_dir) and not os.listdir(fail_dir):
        os.rmdir(fail_dir)
    print("✅ Done.")

if __name__ == "__main__":
    main()
