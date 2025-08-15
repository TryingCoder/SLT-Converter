import os
import sys
import subprocess
import tempfile
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from .utils import validate_file_format

# -----------------------------
# Ensure tqdm installed
# -----------------------------
def ensure_tqdm():
    try:
        import tqdm
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "tqdm"])
        import tqdm
    return tqdm

tqdm = ensure_tqdm()

SOFFICE_PATH = None

def find_soffice():
    import shutil
    paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ]
    for p in paths:
        if os.path.isfile(p):
            return p
    # Try in PATH
    soffice_in_path = shutil.which("soffice")
    if soffice_in_path:
        return soffice_in_path
    return None

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def copy_files(src, dst, ext=".qpw", desc="Copying"):
    files = [f for f in os.listdir(src) if f.lower().endswith(ext.lower())]
    ensure_dir(dst)
    with tqdm.tqdm(total=len(files), desc=desc, unit="file") as pbar:
        for f in files:
            shutil.copy2(os.path.join(src, f), os.path.join(dst, f))
            pbar.update(1)

def convert_file(input_file, out_dir):
    output_file = os.path.join(out_dir, os.path.basename(input_file).replace(".qpw", ".xlsx"))
    if os.path.exists(output_file):
        return True
    try:
        subprocess.run([SOFFICE_PATH, "--headless", "--convert-to", "xlsx", input_file, "--outdir", out_dir],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return os.path.exists(output_file)
    except Exception:
        return False

def convert_all(src_folder, dest_folder, failed_folder, max_workers=4):
    ensure_dir(failed_folder)
    all_files = [f for f in os.listdir(src_folder) if f.lower().endswith(".qpw")]
    pending = set(all_files)
    succeeded = set()
    last_pending = len(pending)

    with tqdm.tqdm(total=len(all_files), desc="Converting", unit="file") as pbar:
        attempt = 1
        while pending:
            failed_this_round = set()
            current_files = list(pending)
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(convert_file, os.path.join(src_folder, f), dest_folder): f for f in current_files}
                for fut in as_completed(futures):
                    f = futures[fut]
                    success = fut.result()
                    if success:
                        if f not in succeeded:
                            pbar.update(1)
                            succeeded.add(f)
                        pending.discard(f)
                        # Remove from failed folder
                        failed_path = os.path.join(failed_folder, f)
                        if os.path.exists(failed_path):
                            os.remove(failed_path)
                    else:
                        failed_this_round.add(f)
                        src_file = os.path.join(src_folder, f)
                        dst_file = os.path.join(failed_folder, f)
                        if not os.path.exists(dst_file):
                            shutil.copy2(src_file, dst_file)

            if not failed_this_round or len(failed_this_round) == last_pending:
                break
            pending = failed_this_round
            last_pending = len(pending)
            attempt += 1

    return succeeded, pending

def main():
    global SOFFICE_PATH

    # Handle --update first
    if len(sys.argv) > 1 and sys.argv[1] == "--update":
        print("Updating SLT-Converter from GitHub...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade",
                               "git+https://github.com/TryingCoder/SLT-Converter.git"])
        print("Update complete.")
        return

    SOFFICE_PATH = find_soffice()
    if not SOFFICE_PATH:
        print("LibreOffice CLI tools not found. Please install 'soffice.exe'.")
        return

    # Source folder input
    src_folder = input("\nEnter SOURCE folder path: ").strip().replace('"', '').replace("'", "").replace("\\", "/")
    while not os.path.isdir(src_folder):
        src_folder = input("Invalid path. Enter SOURCE folder path: ").strip().replace("\\", "/")

    # Destination folder input
    dest_folder = input("Enter DESTINATION folder path: ").strip().replace("\\", "/")
    ensure_dir(dest_folder)

    failed_folder = os.path.join(dest_folder, "Failed")
    ensure_dir(failed_folder)

    max_workers = int(sys.argv[1]) if len(sys.argv) > 1 else 4

    # Temporary working dir
    tmp_dir = tempfile.mkdtemp(prefix="qpw_work_")
    print(f"Temp dir: {tmp_dir}")
    copy_files(src_folder, tmp_dir, desc="Populating temp")

    # Convert all
    succeeded, failed = convert_all(tmp_dir, dest_folder, failed_folder, max_workers)

    shutil.rmtree(tmp_dir)
    print(f"\nConversion complete. {len(succeeded)} succeeded, {len(failed)} failed.")

if __name__ == "__main__":
    main()
