import os
import sys
import ctypes
import subprocess

def get_scripts_path():
    """Return the user Scripts folder path."""
    scripts_path = os.path.join(sys.prefix, "Scripts")
    return scripts_path

def is_in_path(path):
    """Check if a folder is already in PATH."""
    current_paths = os.environ.get("PATH", "").split(os.pathsep)
    return any(os.path.normcase(path) == os.path.normcase(p) for p in current_paths)

def add_to_user_path(path):
    """Add a folder to the user PATH via registry (requires admin for system PATH)."""
    try:
        # Use setx to modify user PATH
        subprocess.run(
            ['setx', 'PATH', f"%PATH%;{path}"],
            shell=True,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        print(f"\n✅ Added {path} to your user PATH. Restart terminal to apply changes.")
    except Exception as e:
        print(f"\n❌ Failed to add PATH automatically. Please add {path} to PATH manually.")
        print(f"Error: {e}")

def add_scripts_to_path():
    """Check and optionally add Python Scripts folder to PATH on Windows."""
    if os.name != "nt":
        return  # Only for Windows

    scripts_path = get_scripts_path()
    if is_in_path(scripts_path):
        return  # Already in PATH

    print(f"\nThe Python Scripts folder is not in your PATH:\n{scripts_path}")
    choice = input("Do you want to add it now? (Y/N): ").strip().lower()
    if choice in ["y", "yes"]:
        # Check for admin privileges
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            is_admin = False

        if is_admin:
            add_to_user_path(scripts_path)
        else:
            print("\n⚠ You need to run this script with administrator privileges to auto-add PATH.")
            print("You can add the following folder manually to your PATH:")
            print(f"  {scripts_path}")
    else:
        print("\n⚠ PATH not modified. You need to add the Scripts folder manually to use CLI commands directly.")

if __name__ == "__main__":
    add_scripts_to_path()
