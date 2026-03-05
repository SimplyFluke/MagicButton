import re
import pyperclip
import pyautogui # pywin32

from time import sleep # For tests
from windows_toasts import Toast, WindowsToaster
from win32gui import GetWindowText, GetForegroundWindow

window = GetWindowText(GetForegroundWindow())

shortcuts = {
    }

# Display toast notification
def toast(message):
    toaster = WindowsToaster("Magic Button")
    newToast = Toast()
    newToast.text_fields = [message]
    toaster.show_toast(newToast)


# Format MAC-address from ugly string
def convertMac(macAddress):
    pattern = re.compile(r"^[0-9a-fA-F]{12}$")
    if pattern.match(macAddress):
        formattedMac = ":".join([macAddress[i:i+2] for i in range(0, 12, 2)])
        pyperclip.copy(formattedMac)
        toast("Converted string to MAC Address.")
    
    else:
        return


# Get printerinfo ezpz
def compilePrinterInfo(printerInfo):
    if len(clip.strip().replace("\r", "").split("\n")) > 1000:
        dump = clip.strip().replace("\r", "").split("\n")[20:60]
    
    else:
        dump = clip.strip().replace("\r", "").split("\n")

    lopenummer = dump[dump.index("Physical identifier")+1]
    lopenummer = re.findall(r"TKP\d+", lopenummer)[0]
    locationInfo = dump[dump.index("Location/Department")+1].split(",")
    rom = dump[dump.index("Location/Department")+1].split(",")[2:5] # Fucking comma
    tmp = ""
    
    for item in rom:
        item = item.lstrip()
        tmp += f"{item},"

    pyperclip.copy(f"Enhet: {locationInfo[0]}\nAdresse: {locationInfo[1]}\nEtg./Rom: {tmp[:-1].capitalize()}\nLøpenummer: {lopenummer}\nModell: {dump[dump.index('Type/Model')+1]}")
    toast("Compiled printer info.")

# Make text lowercase and capitalize after every period
def capitalizeString(text):
    string = re.sub(r'\. *([a-z])', lambda match: '. ' + match.group(1).upper(), text.lower())
    pyperclip.copy(string)
    toast("De-Narinified text.")

clip = pyperclip.paste()

if clip in shortcuts.keys():
    pyperclip.copy (shortcuts[clip])
    toast("Copied text")
    exit()

if "papercut" in window.lower():
    compilePrinterInfo(clip)
    exit()

if len(clip) == 12:
    convertMac(clip)
    exit()

if len(clip) != 12 and clip.isupper(): # Maybe add a minimum requirement for length?
    capitalizeString(clip)
    exit()
