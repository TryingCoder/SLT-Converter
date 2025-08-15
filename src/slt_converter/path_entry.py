import os
import sys
import subprocess

def get_scripts_path():
    """Get he user-specific Python Scripts folder path."""
    user_base = os.path.expanduser("~")
    # Get Python version
    py_version = f"Python{sys.version_info.major}{sys.version_info.minor}"
    scripts_path = os.path.join(user_base, "AppData", "Roaming", "Python", py_version, "Scripts")
    return scripts_path

def is_in_path(path):
    """Check if a given path is in the current PATH environment variable."""
    current_paths = os.environ.get("PATH", "").split(os.pathsep)
    return any(os.path.normcase(path) == os.path.normcase(p) for p in current_paths)

def add_to_path_prompt():
    """Prompt user to add Scripts folder to PATH permanently."""
    scripts_path = get_scripts_path()
    if not os.path.exists(scripts_path):
        print(f"Scripts folder does not exist: {scripts_path}")
        print("You may need to run `pip install --user` first to create it.")
        return

    if is_in_path(scripts_path):
        return  # Already in PATH

    print(f"\nPython Scripts folder not found in PATH:\n  {scripts_path}")
    choice = input("Add slt-converter to PATH? (Y)es/(N)o: ").strip().lower()
    if choice not in ["y", "yes"]:
        print("Skipping PATH update. You will need to run the CLI using full path or add it manually.")
        return

    try:
        # Use setx to add to user PATH
        current_path = os.environ.get("PATH", "")
        new_path = f"{scripts_path};{current_path}"
        subprocess.run(f'setx PATH "{new_path}"', shell=True, check=True)
        print("\nScripts folder added to user PATH. You may need to restart your terminal.")
    except Exception as e:
        print(f"Failed to update PATH: {e}")

def ensure_path():
    if os.name == "nt":
        add_to_path_prompt()
