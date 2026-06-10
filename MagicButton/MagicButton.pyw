import os
import sys
import subprocess
import urllib.request
import webbrowser

def _bootstrap():
    packages = {
        "pyperclip":      "pyperclip",
        "win32gui":       "pywin32",
        "pyautogui":      "pyautogui",
        "windows_toasts": "windows-toasts",
        "pystray":        "pystray",
        "PIL":            "Pillow",
    }
    for import_name, pkg_name in packages.items():
        try:
            __import__(import_name)
        except ImportError:
            print(f"Installing {pkg_name}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name, "--quiet"])
            except subprocess.CalledProcessError:
                print(f"Failed to install {pkg_name}. Try: pip install {pkg_name}")
                sys.exit(1)

_bootstrap()

def _pull_updates():
    import json
    base = "https://raw.githubusercontent.com/SimplyFluke/MagicButton/main/MagicButton/"
    here = os.path.dirname(os.path.abspath(__file__))
    updated = False
    errors = []

    for name in ("MBfunctions.py", "MagicButton.pyw", "woosh64.png"):
        try:
            remote = urllib.request.urlopen(base + name, timeout=5).read()
            local_path = os.path.join(here, name)
            if not os.path.exists(local_path):
                with open(local_path, "wb") as f:
                    f.write(remote)
                updated = True
            else:
                with open(local_path, "rb") as f:
                    if f.read() != remote:
                        with open(local_path, "wb") as f:
                            f.write(remote)
                        updated = True
        except Exception as e:
            errors.append(f"Update failed ({name}): {e}")

    if errors and not updated:
        return "\n".join(errors)

    if not updated:
        return None

    try:
        data = urllib.request.urlopen(base + "changelog.json", timeout=5).read()
        entry = json.loads(data)
        return f"Updated to v{entry['version']}\n{entry['changes']}"
    except Exception:
        return "Magic Button was updated."

_update_message = _pull_updates()
_update_message = None

import re, importlib, threading, pyperclip
import tkinter as tk
from tkinter import messagebox

from time import sleep
from win32gui import GetWindowText, GetForegroundWindow
import win32event, win32api, winerror

import pystray
from PIL import Image

try:
    import MBfunctions as mb
except Exception as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("MagicButton – Load Error", f"Failed to load MBfunctions:\n\n{e}")
    root.destroy()
    sys.exit(1)

if _update_message:
    mb.toast(_update_message)

try: # Create an empty Shortcuts.py with necessary dicts if missing
    from Shortcuts import shortcuts, toastShortcuts

except ModuleNotFoundError:
    with open ("Shortcuts.py", "w") as f:
        f.write("shortcuts = {}\ntoastShortcuts = {}")

    import Shortcuts
    importlib.reload(Shortcuts)
    shortcuts = Shortcuts.shortcuts
    toastShortcuts = Shortcuts.toastShortcuts

__version__ = "3.2.0"

HERE = os.path.dirname(os.path.abspath(__file__))
ICON_PATH = os.path.join(HERE, "woosh64.png")
HELP_URL = "https://docs.google.com/document/d/14rv8u1-lMZvch7fxQTySQ5Knj6ndoKpNErN6Z2zIyb8/edit?tab=t.0"

window = GetWindowText(GetForegroundWindow())
options = []
macPattern = re.compile(r"^[0-9a-fA-F]{12}$")
clip = pyperclip.paste().strip().lstrip()

def ask_user_to_choose_menu(options):
    selected = {"func": None}

    root = tk.Tk()
    root.withdraw()

    menu = tk.Menu(root, tearoff=0)

    def choose(func):
        selected["func"] = func
        root.quit()

    for label, func in options:
        menu.add_command(label=label, command=lambda f=func: choose(f))

    if options:
        menu.add_separator()
    menu.add_command(label="Cancel", command=root.quit)

    x = root.winfo_pointerx()
    y = root.winfo_pointery()

    try:
        menu.tk_popup(x, y)
        root.mainloop()
    finally:
        menu.grab_release()
        root.destroy()

    return selected["func"]

def choose_and_close(root, selected, func):
    selected["func"] = func
    root.destroy()


def run_clip_actions():
    if clip in shortcuts.keys(): # Run shortcuts outside of other functions? Smash through, no fucks?
        pyperclip.copy (shortcuts[clip])
        try:
            mb.toast(f"Copypasta - {toastShortcuts[clip]}")

        except:
            mb.toast("Copied shortcut.")
        return

    if "tkp" in clip.lower() and len(clip) > 10:
        options.append(("Get printer info", lambda: mb.compilePrinterInfo(clip)))

    if "tka" in clip.lower() and len(clip) == 10:
        options.append(("Convert ID from AD", lambda: mb.convertIDfromAD(clip)))
        options.append(("Convert and get info", lambda: mb.convertAndGetInfo(clip)))

    if clip.lower().startswith("tk") and len(clip) < 15 and not clip.lower().startswith("tka"):
        options.append(("Get computer info", lambda: mb.getComputerInfo(clip)))
        options.append(("Get model name only", lambda: mb.getComputerModel(clip)))

    if re.match(r'^KB\d+$', clip, re.IGNORECASE) or ("service-now.com" in clip and ("sysparm_article=" in clip.lower() or "sys_kb_id=" in clip.lower())):
        options.append(("Get KB article title", lambda: mb.getKBTitle(clip)))

    if macPattern.match(clip):
        options.append(("Convert MAC", lambda: mb.convertMac(clip)))

    if len(clip) != 12 and clip.isupper():
        options.append(("Make string lower caps", lambda: mb.capitalizeString(clip)))

    if "google sheets" in window.lower() or "google regneark" in window.lower():
        options.append(("Format sheet (ADRL)", lambda: mb.formatSheetADRL()))

    if "beyondtrust" in window.lower():
        options.append(("G8 freeze link", lambda: mb.g8FreezeLink()))

    if clip.lower() == "shortcuts":
        mb.showShortcuts(shortcuts, toastShortcuts)
        return

    if clip.lower() == "help":
        webbrowser.open(HELP_URL)
        mb.toast("Opening help document...")
        return

    if clip.lower() in ("suggestion", "forslag"):
        mb.sendSuggestion()
        return

    if not options:
        mb.toast(f'Could not find function "{clip}"')

    elif len(options) == 1:
        options[0][1]()

    else:
        chosen = ask_user_to_choose_menu(options)
        if chosen:
            chosen()


run_clip_actions()


# --- System tray icon: only the first invocation stays running ---
mutex = win32event.CreateMutex(None, False, "MagicButtonTrayMutex")

if win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS:

    def on_shortcuts(icon, item):
        threading.Thread(target=mb.showShortcuts, args=(shortcuts, toastShortcuts), daemon=True).start()

    def on_help(icon, item):
        webbrowser.open(HELP_URL)

    def on_check_updates(icon, item):
        def _check():
            message = _pull_updates()
            mb.toast(message or "Magic Button is up to date.")
        threading.Thread(target=_check, daemon=True).start()

    def on_suggestion(icon, item):
        threading.Thread(target=mb.sendSuggestion, daemon=True).start()

    def on_exit(icon, item):
        icon.stop()

    tray_icon = pystray.Icon(
        "MagicButton",
        Image.open(ICON_PATH),
        f"Magic Button v{__version__}",
        menu=pystray.Menu(
            pystray.MenuItem("Shortcuts", on_shortcuts),
            pystray.MenuItem("Help", on_help),
            pystray.MenuItem("Send suggestion", on_suggestion),
            pystray.MenuItem("Check for updates", on_check_updates),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", on_exit),
        ),
    )
    tray_icon.run()
