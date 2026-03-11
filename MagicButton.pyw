import re
import sqlite3
import pyperclip

from windows_toasts import Toast, WindowsToaster
from win32gui import GetWindowText, GetForegroundWindow

window = GetWindowText(GetForegroundWindow())

db = r'C:\ProgramData\TK\Klientoversikt\Computers.SQLite'
intune_url = r'https://intune.microsoft.com/?l=en.en-gb#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId'

shortcuts = {}

toastShortcuts = {} # Aliases for toast notifications when using shortcuts

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
    pyperclip.copy(string.capitalize())
    toast("De-Narinified text.")

# Convert old ID to new
def convertIDfromAD(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            convert = cursor.execute(f"SELECT Computername_Intune, IntuneDeviceID FROM Computers WHERE Computername_AD = '{deviceID}'")
            convert = convert.fetchall()
            pyperclip.copy(convert[0][0])
            
            # Custom toast for computer lookup
            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Converted to new Device ID.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{convert[0][1]}" # Add clickable link to Intune device

            toaster.show_toast(newToast)            
            return

    except Exception as e:
        print (e)
        toast ("Encountered error. Aborting.")
        return

# Get general info about computer
def getComputerInfo(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Model, Enrollment_date, Department, Aktiv, IntuneDeviceID FROM Computers WHERE Computername_Intune = '{deviceID}'")
            info = info.fetchall()
            pyperclip.copy (f"Løpenummer: {deviceID}\nEnhet: {info[0][2]}\nModell: {info[0][0]}\nInnrullert: {info[0][1]}\nAktiv: {info[0][3]}")

            print(f"{intune_url}/{info[0][4]}")
            
            # Custom toast for computer lookup
            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Grabbed computer info.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{info[0][4]}" # Add clickable link to Intune device

            toaster.show_toast(newToast)   
            return

    except:
        toast("Encountered error. Aborting.")
        return
    

clip = pyperclip.paste().strip().lstrip()

if clip in shortcuts.keys():
    pyperclip.copy (shortcuts[clip])
    try:
        toast(f"Copypasta - {toastShortcuts[clip]}")
    except:
        toast("Copied text from shortcut.")

elif "papercut" in window.lower():
    compilePrinterInfo(clip)

elif "tka" in clip.lower() and len(clip) == 10:
    convertIDfromAD(clip)

elif "tk5" in clip.lower() and len(clip) < 15 or "tk8" in clip.lower() and len(clip) < 15 or "tk1" in clip.lower() and len(clip) < 15:
    getComputerInfo(clip)

elif len(clip) == 12:
    convertMac(clip)

elif len(clip) != 12 and clip.isupper(): # Maybe add a minimum requirement for length?
    capitalizeString(clip)

else:
    toast(f'Could not find function "{clip}"')
    


