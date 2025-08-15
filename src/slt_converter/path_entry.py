import os
import sys
import shutil
import subprocess
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def exe_in_path(exe_name):
    return shutil.which(exe_name) is not None

def add_to_user_path(path):
    subprocess.run([
        "setx", "PATH", f"%PATH%;{path}"
    ], shell=True)
    print(f"[+] Added {path} to your USER PATH. Restart your terminal.")

def add_to_system_path(path):
    subprocess.run([
        "powershell",
        "-Command",
        f'Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment" '
        f'-Name PATH -Value ((Get-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment" '
        f'-Name PATH).Path + ";{path}")'
    ], shell=True)
    print(f"[+] Added {path} to your SYSTEM PATH. You may need to log out and back in.")

def main():
    if os.name != "nt":  # Only run on Windows
        return
    
    script_dir = os.path.join(os.path.dirname(sys.executable), "Scripts")
    cli_exe = "slt-converter.exe"  # The entry point pip creates
    
    if exe_in_path(cli_exe):
        return  # Already in PATH
    
    print(f"[!] The CLI command '{cli_exe}' is not in your PATH.")
    choice = input("(Y) Add for this user only (A) Add system-wide (N) No: ").strip().lower()
    
    if choice == "y":
        add_to_user_path(script_dir)
    elif choice == "a":
        if is_admin():
            add_to_system_path(script_dir)
        else:
            print("[!] You must run this with admin rights for system PATH. Falling back to user PATH.")
            add_to_user_path(script_dir)
    else:
        print("[i] Skipping PATH modification. Some features may not work as expected.")

if __name__ == "__main__":
    main()
