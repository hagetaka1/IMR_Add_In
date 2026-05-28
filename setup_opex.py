"""
OPEX IMR Chart - Post-install setup script.
Run automatically by the NSIS installer after files are copied.
Handles: xlwings install, Excel add-in registration, success message.
"""
import sys
import os
import subprocess
import winreg
import tkinter as tk
from tkinter import messagebox


ADDIN_FOLDER = os.path.join(os.environ.get("LOCALAPPDATA", "C:\\"), "OPEX_IMR")
XLA_PATH     = os.path.join(ADDIN_FOLDER, "OPEX.xla")


def show_progress(msg):
    print(msg, flush=True)


def install_xlwings():
    """Install xlwings via pip, return (success, message)."""
    try:
        show_progress("Installing xlwings...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "xlwings", "--quiet"],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0:
            return True, "xlwings installed successfully."
        else:
            return False, f"pip error:\n{result.stderr[:500]}"
    except subprocess.TimeoutExpired:
        return False, "pip timed out after 2 minutes. Check your internet connection."
    except Exception as e:
        return False, str(e)


def register_addin():
    """Register OPEX.xla in the Excel OPEN keys in the registry."""
    try:
        # Find the Excel version key
        excel_key_path = r"Software\Microsoft\Office"
        versions = []
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, excel_key_path) as key:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(key, i)
                    try:
                        float(sub)   # only version numbers like "16.0"
                        versions.append(sub)
                    except ValueError:
                        pass
                    i += 1
                except OSError:
                    break

        if not versions:
            return False, "Could not find Microsoft Office in the registry."

        version = sorted(versions, key=float)[-1]   # use latest version
        options_path = rf"Software\Microsoft\Office\{version}\Excel\Options"

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, options_path,
                            0, winreg.KEY_READ | winreg.KEY_WRITE) as key:

            # Find first free OPEN / OPENn slot, skip if already registered
            existing = []
            try:
                existing.append(winreg.QueryValueEx(key, "OPEN")[0])
            except FileNotFoundError:
                pass
            n = 1
            while True:
                try:
                    existing.append(winreg.QueryValueEx(key, f"OPEN{n}")[0])
                    n += 1
                except FileNotFoundError:
                    break

            # Already registered?
            for val in existing:
                if "OPEX.xla" in val:
                    return True, "Add-in already registered."

            # Register in next free slot
            slot = "OPEN" if not existing else f"OPEN{len(existing)}"
            winreg.SetValueEx(key, slot, 0, winreg.REG_SZ, XLA_PATH)

        return True, f"Add-in registered in Excel {version} (key: {slot})."

    except PermissionError:
        return False, "Permission denied writing to registry. Try running as administrator."
    except Exception as e:
        return False, str(e)


def main():
    root = tk.Tk()
    root.withdraw()

    steps = []

    # Step 1: xlwings
    ok, msg = install_xlwings()
    steps.append(("xlwings install", ok, msg))

    # Step 2: register add-in
    ok2, msg2 = register_addin()
    steps.append(("Excel add-in registration", ok2, msg2))

    all_ok = all(s[1] for s in steps)

    if all_ok:
        messagebox.showinfo(
            "OPEX IMR Chart - Setup Complete",
            "Installation successful!\n\n"
            "The IMR Chart button has been added to your Excel\n"
            "Quick Access Toolbar.\n\n"
            "Restart Excel to activate the add-in."
        )
    else:
        details = "\n\n".join(
            f"{'OK' if ok else 'FAILED'}: {name}\n{msg}"
            for name, ok, msg in steps
        )
        messagebox.showerror(
            "OPEX IMR Chart - Setup Issues",
            f"Setup completed with issues:\n\n{details}\n\n"
            "You can manually install the add-in:\n"
            "Excel -> File -> Options -> Add-Ins -> Go -> Browse\n"
            f"-> {XLA_PATH}"
        )


if __name__ == "__main__":
    main()
