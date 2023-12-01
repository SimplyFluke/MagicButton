import re
import pyperclip
import pyautogui # pywin32

from time import sleep # For tests
from win32gui import GetWindowText, GetForegroundWindow

window = GetWindowText(GetForegroundWindow())

if "papercut" in window.lower():
    pyautogui.hotkey("ctrl", "c") # Remove line if you want to copy yourself
    clip = pyperclip.paste()
    
    if len(clip.strip().replace("\r", "").split("\n")) > 1000: # In case ctrl+a
        dump = clip.strip().replace("\r", "").split("\n")[20:60]
    else:
        dump = clip.strip().replace("\r", "").split("\n")

    lopenummer = dump[dump.index("Physical identifier")+1]
    lopenummer = re.findall("TKP\d+", lopenummer)[0]
    locationInfo = dump[dump.index("Location/Department")+1].split(",")
    rom = dump[dump.index("Location/Department")+1].split(",")[2:5] # Fucking comma
    tmp = ""
    
    for item in rom:
        item = item.lstrip()
        tmp += f"{item},"

    pyperclip.copy(f"Enhet: {locationInfo[0]}\nAdresse: {locationInfo[1]}\nEtg./Rom: {tmp[:-1].capitalize()}\nLøpenummer: {lopenummer}\nModell: {dump[dump.index('Type/Model')+1]}")
