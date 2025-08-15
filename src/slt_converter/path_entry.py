import os
import sys
import site

def add_to_path():
    if os.name != "nt":
        return

    scripts = os.path.join(site.getusersitepackages(), "..", "Scripts")
    scripts = os.path.abspath(scripts)

    current_path = os.environ.get("PATH", "")
    if scripts.lower() in [p.lower() for p in current_path.split(os.pathsep)]:
        return

    print(f"\nYour Python Scripts folder is not in PATH:\n  {scripts}")
    choice = input("Add to PATH now? (Y/N): ").strip().lower()
    if choice in ("y", "yes"):
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_ALL_ACCESS) as env:
                new_path = f"{current_path};{scripts}"
                winreg.SetValueEx(env, "PATH", 0, winreg.REG_EXPAND_SZ, new_path)
            print("PATH updated. You may need to restart your terminal.")
        except Exception as e:
            print("Failed to update PATH automatically. Add manually:", scripts)
