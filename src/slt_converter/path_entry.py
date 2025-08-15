import os
import sys
import site

def add_to_path():
    """Add the user Scripts folder to PATH on Windows if missing."""
    if os.name != "nt":
        return False  # Only applies to Windows

    scripts = os.path.abspath(os.path.join(site.getusersitepackages(), "..", "Scripts"))
    current_path = os.environ.get("PATH", "")

    # Check if Scripts folder is already in PATH (case-insensitive)
    if scripts.lower() in [p.lower() for p in current_path.split(os.pathsep)]:
        return True

    print(f"\nYour Python Scripts folder is not in PATH:\n  {scripts}")
    choice = input("Add to PATH now? (Y/N): ").strip().lower()
    if choice not in ("y", "yes"):
        print("Skipping PATH update.")
        return False

    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_ALL_ACCESS) as env:
            new_path = current_path + os.pathsep + scripts
            winreg.SetValueEx(env, "PATH", 0, winreg.REG_EXPAND_SZ, new_path)
        print("PATH updated successfully. You may need to restart your terminal.")
        return True
    except Exception as e:
        print("Failed to update PATH automatically. Add manually:", scripts)
        return False
