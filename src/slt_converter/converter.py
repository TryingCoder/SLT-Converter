import os
import re
import shutil
import subprocess
import sys
import tempfile
import importlib.util
import site
import argparse
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# -----------------------------
# Version (Beta)
# -----------------------------
__version__ = "0.1.1"

# -----------------------------
# Ensure tqdm installed
# -----------------------------
def ensure_tqdm_installed():
    if importlib.util.find_spec("tqdm") is None:
        print("Installing required package: tqdm...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--no-warn-script-location", "tqdm"])
    user_site = site.getusersitepackages()
    if user_site not in sys.path and os.path.exists(user_site):
        sys.path.append(user_site)
    global tqdm
    from tqdm import tqdm

ensure_tqdm_installed()

# -----------------------------
# Utility Functions
# -----------------------------
def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def copy_files(src_dir, dst_dir, ext=".qpw", progress_desc=None):
    files = [f for f in os.listdir(src_dir) if f.lower().endswith(ext.lower())]
    ensure_dir(dst_dir)
    with tqdm(total=len(files), desc=progress_desc or "Copying files", unit="file") as pbar:
        for filename in files:
            shutil.copy2(os.path.join(src_dir, filename), os.path.join(dst_dir, filename))
            pbar.update(1)

def find_duplicates(folder, extension=".qpw"):
    files = [f for f in os.listdir(folder) if f.lower().endswith(extension.lower())]
    base_map = {}
    for filename in files:
        match = re.match(r"^(.*?)(?: \((\d+)\))?{}$".format(re.escape(extension)), filename, re.IGNORECASE)
        if match:
            base = match.group(1).lower()
            num = int(match.group(2)) if match.group(2) else 0
            base_map.setdefault(base, []).append((num, filename))
    duplicates_to_remove = []
    for base, versions in base_map.items():
        versions.sort()
        for _, filename in versions[1:]:
            duplicates_to_remove.append(filename)
    return duplicates_to_remove

def cleanup_duplicate_files(folder, extension=".qpw", prompt_user=True):
    duplicates_to_remove = find_duplicates(folder, extension)
    if duplicates_to_remove:
        if prompt_user:
            print("\nDuplicate files found:\n")
            for filename in duplicates_to_remove:
                print(f" - {filename}")
            if input("\nRemove duplicate files? (Y)es/(N)o: ").strip().lower() in ["y", "yes"]:
                with tqdm(total=len(duplicates_to_remove), desc="Removing Duplicates", unit="file") as pbar:
                    for filename in duplicates_to_remove:
                        os.remove(os.path.join(folder, filename))
                        pbar.update(1)
                return len(duplicates_to_remove)
            else:
                print("\nNo duplicates removed.")
                return 0
        else:
            return 0
    else:
        if prompt_user:
            print("\nNo duplicate files found.")
        return 0

# -----------------------------
# LibreOffice detection & installation
# -----------------------------
def find_soffice():
    """Cross-platform LibreOffice soffice detection."""
    possible_paths = []
    if sys.platform == "win32":
        possible_paths = [
            os.environ.get("PROGRAMFILES", r"C:\Program Files") + r"\LibreOffice\program\soffice.exe",
            os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)") + r"\LibreOffice\program\soffice.exe",
        ]
    elif sys.platform == "darwin":
        possible_paths = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    else:
        possible_paths = ["/usr/bin/soffice", "/usr/local/bin/soffice"]

    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None

def get_latest_libreoffice_windows_url():
    """Scrape LibreOffice download page to find latest Windows x86_64 MSI URL."""
    base_url = "https://download.documentfoundation.org/libreoffice/stable/"
    resp = requests.get(base_url)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    versions = [a.text.strip("/") for a in soup.find_all("a") if a.text[0].isdigit()]
    latest_version = sorted(versions, key=lambda s: list(map(int, s.split("."))))[-1]
    return f"{base_url}{latest_version}/win/x86_64/LibreOffice_{latest_version}_Win_x64.msi"

def install_libreoffice():
    """Install LibreOffice CLI tools headlessly (Windows or Linux)."""
    print("\nLibreOffice CLI not found. Installing...")
    if sys.platform == "win32":
        try:
            url = get_latest_libreoffice_windows_url()
        except Exception as e:
            print(f"❌ Could not fetch latest LibreOffice version: {e}")
            return
        installer = os.path.join(tempfile.gettempdir(), "LibreOffice.msi")
        subprocess.run(["powershell", "-Command", f"Invoke-WebRequest -Uri {url} -OutFile {installer}"], check=True)
        print("Installing LibreOffice headlessly...")
        subprocess.run(["msiexec", "/i", installer, "/quiet", "/norestart"], check=True)
        os.remove(installer)
    elif sys.platform.startswith("linux"):
        if shutil.which("apt"):
            subprocess.run(["sudo", "apt", "update"], check=True)
            subprocess.run(["sudo", "apt", "install", "-y", "libreoffice"], check=True)
        elif shutil.which("yum"):
            subprocess.run(["sudo", "yum", "install", "-y", "libreoffice"], check=True)
        else:
            print("Unsupported Linux distro. Install LibreOffice manually.")
    else:
        print("Unsupported OS. Install LibreOffice manually.")

# -----------------------------
# Conversion functions
# -----------------------------
def convert_qpw_to_xlsx(input_file, output_dir, soffice_path=None):
    if soffice_path is None:
        soffice_path = find_soffice()
    if soffice_path is None:
        raise FileNotFoundError("LibreOffice soffice executable not found.")
    output_file = os.path.join(output_dir, os.path.basename(input_file).replace('.qpw', '.xlsx'))
    if os.path.exists(output_file):
        return True
    try:
        command = [soffice_path, '--headless', '--convert-to', 'xlsx', input_file, '--outdir', output_dir]
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return os.path.exists(output_file)
    except Exception:
        return False

def convert_all_with_retries(source_folder, converted_folder, failed_folder, max_workers=4, soffice_path=None):
    ensure_dir(failed_folder)
    all_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.qpw')]
    pending_files = set(all_files)
    succeeded_files = set()
    last_pending_count = len(pending_files)

    with tqdm(total=len(all_files), desc="Converting", unit="file") as pbar:
        attempt = 1
        while pending_files:
            current_attempt_files = list(pending_files)
            failed_this_round = set()

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(convert_qpw_to_xlsx, os.path.join(source_folder, f), converted_folder, soffice_path): f for f in current_attempt_files}
                for future in as_completed(futures):
                    filename = futures[future]
                    success = future.result()
                    if success:
                        if filename not in succeeded_files:
                            pbar.update(1)
                            succeeded_files.add(filename)
                            pending_files.discard(filename)
                        failed_path = os.path.join(failed_folder, filename)
                        if os.path.exists(failed_path):
                            os.remove(failed_path)
                    else:
                        failed_this_round.add(filename)
                        src_file = os.path.join(source_folder, filename)
                        dst_file = os.path.join(failed_folder, filename)
                        if not os.path.exists(dst_file):
                            shutil.copy2(src_file, dst_file)
                        with open(os.path.join(failed_folder, "failed_conversions.log"), "a") as log_file:
                            log_file.write(f"{filename}\n")

            if not failed_this_round:
                break
            if len(failed_this_round) == last_pending_count:
                print(f"\nNo further progress after attempt {attempt}, {len(failed_this_round)} files failed.")
                break

            pending_files = failed_this_round
            last_pending_count = len(pending_files)
            attempt += 1

    return succeeded_files, pending_files

# -----------------------------
# Main CLI
# -----------------------------
def main():
    parser = argparse.ArgumentParser(description="Convert QPW files to XLSX using LibreOffice.")
    parser.add_argument("--source", "-s", help="Source folder containing QPW files")
    parser.add_argument("--dest", "-d", help="Destination folder for XLSX files")
    parser.add_argument("--backup", "-b", help="Optional backup folder")
    parser.add_argument("--workers", "-w", type=int, default=4, help="Number of concurrent workers")
    parser.add_argument("--skip-duplicates", action="store_true", help="Skip duplicate file prompt")
    parser.add_argument("--update", action="store_true", help="Update the tool from GitHub")
    parser.add_argument("--version", "-v", action="store_true", help="Show the installed version and exit")
    args = parser.parse_args()

    if args.version:
        print(f"SLT-Converter version {__version__}")
        return

    if args.update:
        print("\nUpdating SLT-Converter from GitHub...\n")
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "git+https://github.com/TryingCoder/SLT-Converter.git"], check=True)
        print("\nUpdate complete!")
        return

    if not args.source or not os.path.isdir(args.source):
        print(f"❌ Source folder does not exist: {args.source}")
        return

    if not args.dest:
        print("❌ Destination folder not specified")
        return
    ensure_dir(args.dest)
    failed_folder = os.path.join(args.dest, "Failed")
    ensure_dir(failed_folder)

    if args.backup:
        ensure_dir(args.backup)
        files_to_backup = [f for f in os.listdir(args.source) if f.lower().endswith(".qpw")]
        with tqdm(total=len(files_to_backup), desc="Creating Backup", unit="file") as pbar:
            for f in files_to_backup:
                shutil.copy2(os.path.join(args.source, f), os.path.join(args.backup, f))
                pbar.update(1)
        print(f"✅ Backup completed at: {args.backup}")

    soffice_path = find_soffice()
    if soffice_path is None:
        choice = input("\nLibreOffice CLI not found. Install now? (Y/N): ").strip().lower()
        if choice in ("y", "yes"):
            install_libreoffice()
            soffice_path = find_soffice()
            if soffice_path is None:
                print("❌ LibreOffice installation failed. Install manually.")
                return
        else:
            print("❌ Install LibreOffice CLI first.")
            return

    working_dir = tempfile.mkdtemp(prefix="qpw_work_")
    try:
        copy_files(args.source, working_dir, ext=".qpw", progress_desc=f"Populating working directory: {working_dir}")
        duplicates_removed = cleanup_duplicate_files(working_dir, extension=".qpw", prompt_user=not args.skip_duplicates)
        print(f"\nConverting {len([f for f in os.listdir(working_dir) if f.lower().endswith('.qpw')])} files...")
        succeeded, failed = convert_all_with_retries(working_dir, args.dest, failed_folder, args.workers, soffice_path)
    finally:
        shutil.rmtree(working_dir)
        print(f"\nTemp working directory removed: {working_dir}")

    total_files = len([f for f in os.listdir(args.source) if f.lower().endswith('.qpw')])
    converted_files = len(succeeded)
    failed_files = len(failed)

    print("\n===== Conversion Summary =====")
    print(f"✅ Total QPW files processed: {total_files}")
    print(f"✅ Successfully converted:    {converted_files}")
    print(f"❌ Failed after all retries:  {failed_files}")
    print(f"✅ Duplicate files removed: {duplicates_removed}")
    print("==============================")

    if not failed and os.path.exists(failed_folder) and not os.listdir(failed_folder):
        os.rmdir(failed_folder)
        print(f"\n✅ No failed conversions. Removed Failed folder: {failed_folder}")

    print("\n✅ All files successfully converted!")

if __name__ == "__main__":
    main()
