import os
import urllib.request

def _pull_updates():
    import json
    base = "https://raw.githubusercontent.com/SimplyFluke/MagicButton/main/MagicButton/"
    here = os.path.dirname(os.path.abspath(__file__))
    updated = False
    error = None

    for name in ("MBfunctions.py", "MagicButton.pyw"):
        try:
            remote = urllib.request.urlopen(base + name, timeout=5).read()
            local_path = os.path.join(here, name)
            with open(local_path, "rb") as f:
                if f.read() != remote:
                    with open(local_path, "wb") as f:
                        f.write(remote)
                    updated = True
        except Exception as e:
            error = f"Update failed ({name}): {e}"

    if error and not updated:
        return error

    if not updated:
        return None

    try:
        data = urllib.request.urlopen(base + "changelog.json", timeout=5).read()
        entry = json.loads(data)
        return f"Updated to v{entry['version']}\n{entry['changes']}"
    except Exception:
        return "Magic Button was updated."

_update_message = _pull_updates()

import re, importlib, pyperclip
import tkinter as tk
import MBfunctions as mb

from time import sleep
from win32gui import GetWindowText, GetForegroundWindow

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

__version__ = "3.0.0"

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


if clip in shortcuts.keys(): # Run shortcuts outside of other functions? Smash through, no fucks?
    pyperclip.copy (shortcuts[clip])
    try:
        mb.toast(f"Copypasta - {toastShortcuts[clip]}")

    except:
        mb.toast("Copied shortcut.")
    exit()

if "|" in clip:
    textInput = clip.split("|")
    url = textInput[0].strip()
    name = textInput[1].strip()
    mb.createSNHyperlink(url, name)

if "tkp" in clip.lower() and len(clip) > 10:
    options.append(("Get printer info", lambda: mb.compilePrinterInfo(clip)))

if "tka" in clip.lower() and len(clip) == 10:
    options.append(("Convert ID from AD", lambda: mb.convertIDfromAD(clip)))
    options.append(("Convert and get info", lambda: mb.convertAndGetInfo(clip)))

if clip.lower().startswith("tk") and len(clip) < 15 and not clip.lower().startswith("tka"):
    options.append(("Get computer info", lambda: mb.getComputerInfo(clip)))
    options.append(("Get model name only", lambda: mb.getComputerModel(clip)))

if macPattern.match(clip):
    options.append(("Convert MAC", lambda: mb.convertMac(clip)))

if len(clip) != 12 and clip.isupper():
    options.append(("Make string lower caps", lambda: mb.capitalizeString(clip)))

if "google sheets" in window.lower() or "google regneark" in window.lower():
    options.append(("Format sheet (ADRL)", lambda: mb.formatSheetADRL()))

if not options:
    mb.toast(f'Could not find function "{clip}"')

elif len(options) == 1:
    options[0][1]()

else:
    chosen = ask_user_to_choose_menu(options)
    if chosen:
        chosen()
