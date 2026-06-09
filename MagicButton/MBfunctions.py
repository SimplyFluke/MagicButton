import os
import re
import sys
import sqlite3
import subprocess
import pyautogui
import pyperclip

from time import sleep
from windows_toasts import Toast, WindowsToaster

db = r'C:\ProgramData\TK\Klientoversikt\Computers.SQLite'
intune_url = r'https://intune.microsoft.com/?l=en.en-gb#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId'
sn_base_url = "https://YOUR_INSTANCE.service-now.com"  # Set your ServiceNow instance URL

def _ensure(package, import_name=None):
    try:
        __import__(import_name or package)
    except ImportError:
        print(f"Installing required package: {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except subprocess.CalledProcessError:
            toast(f"Failed to install {package}.\nTry: pip install {package}")
            return False
    return True

# Display toast notification
def toast(message):
    toaster = WindowsToaster("Magic Button")
    newToast = Toast()
    newToast.text_fields = [message]
    toaster.show_toast(newToast)

# Format MAC-address from Intune's ugly string
def convertMac(macAddress):
    formattedMac = ":".join([macAddress[i:i+2] for i in range(0, 12, 2)])
    pyperclip.copy(formattedMac)
    toast("Converted string to MAC Address.")


# Get printerinfo ezpz
def compilePrinterInfo(printerInfo):
    if len(printerInfo.strip().replace("\r", "").split("\n")) > 1000:
        dump = printerInfo.strip().replace("\r", "").split("\n")[20:60]
    
    else:
        dump = printerInfo.strip().replace("\r", "").split("\n")

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


def convertIDfromAD(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Computername_Intune, IntuneDeviceID FROM Computers WHERE Computername_AD = '{deviceID.upper()}'")
            info = info.fetchall()
            pyperclip.copy(info[0][0])
            
            # Custom toast for computer lookup
            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Converted to new Device ID.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{info[0][1]}"

            toaster.show_toast(newToast)            
            return

    except Exception as e:
        print (e)
        toast ("Encountered error. Aborting.\nDevice might not exist yet.")
        return


def getComputerModel(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Model FROM Computers WHERE Computername_Intune = '{deviceID.upper()}'")
            info = info.fetchall()
            pyperclip.copy(info[0][0])

            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Copied model name"]

            toaster.show_toast(newToast)
    except:
        toast("Encountered error. Aborting.\nDevice might not exist yet.")
        return


def convertAndGetInfo(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Computername_Intune, IntuneDeviceID FROM Computers WHERE Computername_AD = '{deviceID.upper()}'")
            info = info.fetchall()
            info = cursor.execute(f"SELECT Model, Enrollment_date, Department, Aktiv, IntuneDeviceID FROM Computers WHERE Computername_Intune = '{deviceID.upper()}'")
            info = info.fetchall()
            pyperclip.copy (f"Løpenummer: {deviceID.upper()}\nEnhet: {info[0][2]}\nModell: {info[0][0]}\nInnrullert: {info[0][1]}\nAktiv: {info[0][3]}")

            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Converted from old ID and grabbed computer info.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{info[0][4]}"

            toaster.show_toast(newToast)
    except:
        toast("Encountered error. Aborting.\nDevice might not exist yet.")
        return

def g8FreezeLink():
    pyperclip.copy("https://www.intel.com/content/www/us/en/download/864990/intel-11th-14th-gen-processor-graphics-windows.html")
    toast ("Copied G8 freeze link.")

def getComputerInfo(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Model, Enrollment_date, Department, Aktiv, IntuneDeviceID FROM Computers WHERE Computername_Intune = '{deviceID.upper()}'")
            info = info.fetchall()
            pyperclip.copy (f"Løpenummer: {deviceID.upper()}\nEnhet: {info[0][2]}\nModell: {info[0][0]}\nInnrullert: {info[0][1]}\nAktiv: {info[0][3]}")

            print(f"{intune_url}/{info[0][4]}")
            # Custom toast for computer lookup
            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Grabbed computer info.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{info[0][4]}"

            toaster.show_toast(newToast)   
            return

    except:
        toast("Encountered error. Aborting.\nDevice might not exist yet.")
        return
    

def formatSheetADRL():
    # Move to leftmost cell
    pyautogui.leftClick()
    pyautogui.press('home')
    sleep(0.1)

    # Mark leftmost column
    pyautogui.keyDown("ctrl")
    pyautogui.press("space")
    pyautogui.keyUp("ctrl")
    sleep(0.1)
    

    # Split into columns
    pyautogui.keyDown("alt")
    pyautogui.press("d")
    pyautogui.keyUp("alt")
    pyautogui.press("e")
    sleep(0.1)

    # Alternating colors
    pyautogui.keyDown("alt")
    pyautogui.press("o")
    pyautogui.keyUp("alt")
    pyautogui.press("l")
    pyautogui.keyDown("shift")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.keyUp("shift")
    pyautogui.press("enter")
    sleep(0.1)

    # Create filter
    pyautogui.keyDown("alt")
    pyautogui.press("d")
    pyautogui.keyUp("alt")
    pyautogui.press("f")

    toast("Formatted sheet for ADRL")
    return

def showShortcuts(shortcuts, toastShortcuts):
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("Magic Button Shortcuts")
    root.minsize(700, 300)

    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    tree = ttk.Treeview(frame, columns=("trigger", "name", "copies"), show="headings")
    tree.heading("trigger", text="Trigger")
    tree.heading("name", text="Name")
    tree.heading("copies", text="Copies")
    tree.column("trigger", width=180, anchor="w")
    tree.column("name", width=250, anchor="w")
    tree.column("copies", width=420, anchor="w")

    for key, value in shortcuts.items():
        tree.insert("", tk.END, values=(key, toastShortcuts.get(key, ""), value))

    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    detail = tk.Text(root, height=5, wrap=tk.WORD, state=tk.DISABLED, relief=tk.FLAT,
                     bg=root.cget("bg"), padx=10, pady=6)
    detail.pack(fill=tk.X, padx=10, pady=(0, 10))

    def on_select(_):
        selected = tree.focus()
        if not selected:
            return
        copies = tree.item(selected, "values")[2]
        detail.config(state=tk.NORMAL)
        detail.delete("1.0", tk.END)
        detail.insert(tk.END, copies)
        detail.config(state=tk.DISABLED)

    tree.bind("<<TreeviewSelect>>", on_select)
    root.mainloop()


def createSNHyperlink(url, name):
    hyperlink = f'[code]<a href="{url}" target="_blank">{name}</a>[/code]'
    pyperclip.copy(hyperlink)
    toast("Created hyperlink for SN.")


def getKBTitle(kb_input):
    if not _ensure("selenium") or not _ensure("webdriver-manager", "webdriver_manager"):
        return

    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.microsoft import EdgeChromiumDriverManager

    if kb_input.startswith("http"):
        instance_match = re.match(r"https://([\w-]+)\.service-now\.com", kb_input)
        if not instance_match:
            toast("Could not parse ServiceNow URL.")
            return
        base_url = f"https://{instance_match.group(1)}.service-now.com"

        kb_match = re.search(r"sysparm_article=(KB\d+)", kb_input, re.IGNORECASE)
        sys_kb_match = re.search(r"sys_kb_id=([a-f0-9]+)", kb_input, re.IGNORECASE)

        if kb_match:
            kb_number = kb_match.group(1).upper()
        elif sys_kb_match:
            kb_number = None
        else:
            toast("Could not find a KB identifier in the URL.")
            return
    else:
        base_url = sn_base_url
        kb_number = kb_input.upper()

    if kb_number:
        page_url = f"{base_url}/ap?id=kb_article_view&sysparm_article={kb_number}"
    else:
        page_url = kb_input

    profile_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "edge_profile")

    def make_driver(headless):
        options = Options()
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_argument("--profile-directory=Default")
        if headless:
            options.add_argument("--headless=new")
        service = Service(EdgeChromiumDriverManager().install())
        return webdriver.Edge(service=service, options=options)

    def is_login_page(title):
        return not title or "pålogging" in title.lower() or "login" in title.lower()

    def resolve_kb(driver):
        WebDriverWait(driver, 15).until(lambda d: d.title not in ("", "ServiceNow"))
        if is_login_page(driver.title):
            return None, None, None
        h1 = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "h1"))
        )
        title = h1.text.strip() or None

        resolved_kb = kb_number
        if not resolved_kb:
            for src in (driver.current_url, driver.page_source):
                m = re.search(r'\bKB\d+\b', src, re.IGNORECASE)
                if m:
                    resolved_kb = m.group(0).upper()
                    break

        if resolved_kb:
            clean_url = f"{base_url}/ap?id=kb_article_view&sysparm_article={resolved_kb}"
        else:
            clean_url = page_url

        return title, resolved_kb, clean_url

    driver = None
    try:
        driver = make_driver(headless=True)
        driver.get(page_url)
        title, resolved_kb, clean_url = resolve_kb(driver)
        driver.quit()
        driver = None

        if title:
            label = f"{resolved_kb} - {title}" if resolved_kb else title
            createSNHyperlink(clean_url, label)
            return

        # Session expired — open visible window so user can log in
        toast("Session expired.\nPlease log in to ServiceNow in the browser window.")
        driver = make_driver(headless=False)
        driver.get(page_url)
        WebDriverWait(driver, 120).until(lambda d: not is_login_page(d.title))
        title, resolved_kb, clean_url = resolve_kb(driver)
        driver.quit()
        driver = None
        if title:
            label = f"{resolved_kb} - {title}" if resolved_kb else title
            createSNHyperlink(clean_url, label)
        else:
            toast("Logged in but could not find article title.")

    except Exception as e:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        toast(f"Error fetching KB title:\n{e}")
