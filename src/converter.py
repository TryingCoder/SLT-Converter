import os
import re
import shutil
import subprocess
import sys
import tempfile
import importlib.util
import site
from concurrent.futures import ThreadPoolExecutor, as_completed

# -----------------------------
# Ensure tqdm installed
# -----------------------------
def ensure_tqdm_installed():
    if importlib.util.find_spec("tqdm") is None:
        print("Installing required package: tqdm...\n")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--no-warn-script-location", "tqdm"])

    user_site = site.getusersitepackages()
    if user_site not in sys.path and os.path.exists(user_site):
        sys.path.append(user_site)

    global tqdm
    from tqdm import tqdm

ensure_tqdm_installed()

SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

# -----------------------------
# Utility Functions
# -----------------------------

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def copy_files(src_dir, dst_dir, ext=".qpw", progress_desc=None):
    files = [f for f in os.listdir(src_dir) if f.lower().endswith(ext.lower())]
    ensure_dir(dst_dir)
    from tqdm import tqdm
    with tqdm(total=len(files), desc=progress_desc or "Copying files", unit="file") as pbar:
        for filename in files:
            shutil.copy2(os.path.join(src_dir, filename), os.path.join(dst_dir, filename))
            pbar.update(1)

def find_duplicates(folder, extension=".qpw"):
    files = [f for f in os.listdir(folder) if f.lower().endswith(extension.lower())]
    base_map = {}
    for filename in files:
        match = re.match(r"^(.*?)(?: \((\d+)\))?{}".format(re.escape(extension)), filename, re.IGNORECASE)
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
                from tqdm import tqdm
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

def convert_qpw_to_xlsx(input_file, output_dir):
    output_file = os.path.join(output_dir, os.path.basename(input_file).replace('.qpw', '.xlsx'))
    if os.path.exists(output_file):
        return True
    try:
        command = [SOFFICE_PATH, '--headless', '--convert-to', 'xlsx', input_file, '--outdir', output_dir]
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return os.path.exists(output_file)
    except Exception:
        return False

def convert_all_with_retries(source_folder, converted_folder, failed_folder, max_workers=4):
    ensure_dir(failed_folder)
    all_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.qpw')]
    pending_files = set(all_files)
    succeeded_files = set()
    last_pending_count = len(pending_files)

    from tqdm import tqdm
    with tqdm(total=len(all_files), desc="Converting", unit="file") as pbar:
        attempt = 1
        while pending_files:
            current_attempt_files = list(pending_files)
            failed_this_round = set()

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(convert_qpw_to_xlsx, os.path.join(source_folder, f), converted_folder): f for f in current_attempt_files}
                for future in as_completed(futures):
                    filename = futures[future]
                    success = future.result()
                    if success:
                        if filename not in succeeded_files:
                            pbar.update(1)
                            succeeded_files.add(filename)
                            if filename in pending_files:
                                pending_files.remove(filename)
                        # Remove from failed folder if exists
                        failed_path = os.path.join(failed_folder, filename)
                        if os.path.exists(failed_path):
                            os.remove(failed_path)
                    else:
                        failed_this_round.add(filename)
                        # Copy failed file to failed folder if not already there
                        src_file = os.path.join(source_folder, filename)
                        dst_file = os.path.join(failed_folder, filename)
                        if not os.path.exists(dst_file):
                            shutil.copy2(src_file, dst_file)

            if not failed_this_round:
                break  # All done

            if len(failed_this_round) == last_pending_count:
                print(f"\nNo further progress after attempt {attempt}, {len(failed_this_round)} files failed.")
                break

            pending_files = failed_this_round
            last_pending_count = len(pending_files)
            attempt += 1

    return succeeded_files, pending_files

def check_backup_conflicts(source_folder, backup_folder, ext=".qpw"):
    conflicting_files = []
    for filename in os.listdir(source_folder):
        if filename.lower().endswith(ext.lower()):
            backup_file = os.path.join(backup_folder, filename)
            if os.path.exists(backup_file):
                conflicting_files.append(filename)
    if conflicting_files:
        print("\nThe following files already exist in the backup folder:\n")
        for f in conflicting_files:
            print(f" - {f}")
        while True:
            choice = input("\nDo you want to (S)kip or (R)eplace? [S/R]: \n").strip().lower()
            if choice in ['s', 'skip']:
                return "skip"
            elif choice in ['r', 'replace']:
                return "replace"
            else:
                print("\nPlease enter 'S' to skip or 'R' to replace.")
    return "replace"  # Default if no conflicts


def main():
    original_folder = input("\nEnter SOURCE folder path: ").strip()
    while not os.path.isdir(original_folder):
        print("\n❌ Invalid path. Try again.")
        original_folder = input("\nEnter SOURCE folder path: ").strip()

    converted_folder = input("Enter DESTINATION folder path: ").strip()
    ensure_dir(converted_folder)

    failed_folder = os.path.join(converted_folder, "Failed")
    ensure_dir(failed_folder)

    max_workers = int(sys.argv[1]) if len(sys.argv) > 1 else 4

    # Step 1: Backup
    if input("\nCreate a backup of the source files? (Y)es/(N)o: ").strip().lower() == "y":
        backup_folder = input("Enter BACKUP folder path: ").strip()
        ensure_dir(backup_folder)
        action = check_backup_conflicts(original_folder, backup_folder, ext=".qpw")

        all_src_files = [f for f in os.listdir(original_folder) if f.lower().endswith(".qpw")]
        from tqdm import tqdm
        with tqdm(total=len(all_src_files), desc="Creating Backup", unit="file") as pbar:
            for filename in all_src_files:
                src_file = os.path.join(original_folder, filename)
                dst_file = os.path.join(backup_folder, filename)
                if os.path.exists(dst_file):
                    if action == "replace":
                        shutil.copy2(src_file, dst_file)
                    # else skip
                else:
                    shutil.copy2(src_file, dst_file)
                pbar.update(1)
        print(f"\nBackup completed: {backup_folder}")

    # Step 2: Temp working directory
    working_dir = tempfile.mkdtemp(prefix="qpw_work_")
    print(f"\nTemp working directory created: {working_dir}")

    # Step 3: Populate temp working directory
    copy_files(original_folder, working_dir, ext=".qpw", progress_desc="Populating")

    # Step 4: Remove duplicates once, prompt user
    duplicates_removed = cleanup_duplicate_files(working_dir, extension=".qpw", prompt_user=True)

    # Step 5: Convert with retries and single progress bar
    print(f"\nConverting {len([f for f in os.listdir(working_dir) if f.lower().endswith('.qpw')])} files...")
    succeeded, failed = convert_all_with_retries(working_dir, converted_folder, failed_folder, max_workers)

    # Step 6: Cleanup temp folder
    shutil.rmtree(working_dir)
    print(f"\nTemp working directory removed: {working_dir}")

    # Step 7: Summary
    total_files = len([f for f in os.listdir(original_folder) if f.lower().endswith('.qpw')])
    converted_files = len(succeeded)
    failed_files = len(failed)

    print("\n===== Conversion Summary =====")
    print(f"✅ Total QPW files processed: {total_files}")
    print(f"✅ Successfully converted:    {converted_files}")
    print(f"❌ Failed after all retries:  {failed_files}")
    print(f"✅ Duplicate files removed: {duplicates_removed}")
    print("==============================")

    # Step 8: Remove empty Failed folder if no failures
    if not failed and os.path.exists(failed_folder) and not os.listdir(failed_folder):
        try:
            os.rmdir(failed_folder)
            print(f'\n✅ No failed conversions. Removed Failed folder: {failed_folder}')
        except Exception:
            pass

    print("\n✅ All operations complete.")

if __name__ == "__main__":
    main()
