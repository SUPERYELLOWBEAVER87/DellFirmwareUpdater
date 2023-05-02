import selenium
import subprocess
import os
import ctypes
import socket
import base64
import sys
import argparse
import PySimpleGUI as sg
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from random import randint
from time import sleep


print("""\
 _______   ______   ______    ______                                                
|       \ |      \ /      \  /      \                                               
| $$$$$$$\ \$$$$$$|  $$$$$$\|  $$$$$$\                                              
| $$__/ $$  | $$  | $$  | $$| $$___\$$                                              
| $$    $$  | $$  | $$  | $$ \$$    \                                               
| $$$$$$$\  | $$  | $$  | $$ _\$$$$$$\                                              
| $$__/ $$ _| $$_ | $$__/ $$|  \__| $$                                              
| $$    $$|   $$ \ \$$    $$ \$$    $$                                              
 \$$$$$$$  \$$$$$$  \$$$$$$   \$$$$$$                                                                                                                
                                                                                    
 __    __  _______   _______    ______  ________  ________  _______                 
|  \  |  \|       \ |       \  /      \|        \|        \|       \                
| $$  | $$| $$$$$$$\| $$$$$$$\|  $$$$$$\\$$$$$$$$| $$$$$$$$| $$$$$$$\               
| $$  | $$| $$__/ $$| $$  | $$| $$__| $$  | $$   | $$__    | $$__| $$               
| $$  | $$| $$    $$| $$  | $$| $$    $$  | $$   | $$  \   | $$    $$               
| $$  | $$| $$$$$$$ | $$  | $$| $$$$$$$$  | $$   | $$$$$   | $$$$$$$\               
| $$__/ $$| $$      | $$__/ $$| $$  | $$  | $$   | $$_____ | $$  | $$               
 \$$    $$| $$      | $$    $$| $$  | $$  | $$   | $$     \| $$  | $$               
  \$$$$$$  \$$       \$$$$$$$  \$$   \$$   \$$    \$$$$$$$$ \$$   \$$    

  Launch the executable with -h for help and all possible handles.                                                                                
                                                                                    
                    Version: 1.0                                                             
""")

parser = argparse.ArgumentParser(description="Script for automating updating BIOS Firmware on Dell machines.")
parser.add_argument("--v", "-v", help="Makes the browser visible during use of the script. Disables headless mode.", action='store_true')
parser.add_argument("--t","-t", help="For debugging purposes only. Doesn't execute the command to start BIOS firmware update, only shows the string it is going to execute.", action='store_true')
parser.add_argument("--l", "-l", help="For imaging labs. Doesn't ask the user if they want to proceed with the firmware update. Makes the process entirely automated.", action='store_true')
args = parser.parse_args()

if not os.path.exists('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'):
    ctypes.windll.user32.MessageBoxW(0, "Chromedriver requires Google Chrome to be installed!", "Missing Dependency", 0)
    sys.exit(0)

try:
    print("\n[INFO]::Checking user permissions...")
    if ctypes.windll.shell32.IsUserAnAdmin() != 1:
        ctypes.windll.user32.MessageBoxW(0, "Program needs to be run as admin.", "Permission Error", 0)
        sys.exit()
    print("[SUCCESS]::User is an admin.")
except SystemExit:
    sys.exit()
except:
    sys.exit()


val = subprocess.check_output('powershell.exe gwmi win32_bios').decode('utf-8').strip().split("\n")
val2 = subprocess.check_output('manage-bde -status').decode('utf-8').strip().split("\n")
SERIALNUMBER = subprocess.check_output('wmic bios get serialnumber').decode('utf-8').replace('\r','').replace('\n','').replace('SerialNumber','').strip()
BITLOCKERSTATUS = val2[13].replace("Protection Status:", "").strip(" ")
CURRENTBIOSVERSION = val[0].replace('SMBIOSBIOSVersion :','').strip()
HEADLESS = True
TESTMODE = False
LABMODE = False

if args.v:
    HEADLESS = False

if args.t:
    print("[INFO]::DEBUGGING MODE IS ENABLED.")
    TESTMODE = True

if args.l:
    print("[INFO]::LAB MODE IS ENABLED.")
    LABMODE = True

def downloadBIOS():
    global COUNTER

    print("[INFO]::Current BIOS Version: {0}".format(CURRENTBIOSVERSION))

    try:
        options = selenium.webdriver.chrome.options.Options()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument('--window-size=1920,1080')
        options.add_argument("--disable-blink-features=AutomationControlled") 
        options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
        options.add_experimental_option("useAutomationExtension", False) 
        options.add_experimental_option("detach", True)

        if HEADLESS == True:
            options.add_argument('--headless=new')


        driver = selenium.webdriver.Chrome(service=selenium.webdriver.chrome.service.Service(executable_path=ChromeDriverManager().install()), options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") 
        params = {'behavior' : 'allow', 'downloadPath':r"C:\Software"}
        driver.execute_cdp_cmd('Page.setDownloadBehavior', params)

        print("[INFO]::Navigating to Dell Warranty and Contracts...")
        driver.get('https://www.dell.com/support/home/en-us?app=warranty')

        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located((By.ID, "inpEntrySelection")))

        serialQuery = driver.find_element(By.ID, "inpEntrySelection")
        searchBtn = driver.find_element(By.ID, "btn-entry-select")

        print("[INFO]::Sending serial number [{0}] into query.".format(SERIALNUMBER))
        serialQuery.send_keys(SERIALNUMBER)
        sleep(randint(10, 12))

        searchBtn.click()
        print("[INFO]::Searching for serial.")


        try:
            WebDriverWait(driver, 5).until(expected_conditions.presence_of_element_located((By.ID, "iframeSurvey")))
            driver.find_element(By.ID, "iframeSurvey")
            print("[INFO]::Dell Survey exists. Closing survey.")
            driver.switch_to.frame(driver.find_element(By.ID, "iframeSurvey"))
            WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((By.ID, "noButtonIPDell")))
            surveyNo = driver.find_element(By.ID, "noButtonIPDell")
            surveyNo.click()
            driver.switch_to.default_content()
        except :
            print("[INFO]::Did not detect Dell Survey. Continuing.")


        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located((By.ID, "warrantyDetailsHeader")))
        detailsExit = driver.find_element(By.ID, "warranty-cancel")
        element = WebDriverWait(driver, 30).until(expected_conditions.element_to_be_clickable((By.ID, "warranty-cancel")))
        print("[INFO]::Closing asset details pane.")
        detailsExit.click()


        element = WebDriverWait(driver, 30).until(expected_conditions.presence_of_element_located((By.ID, "drivers")))

        driversCat = driver.find_element(By.ID, "drivers")
        element = WebDriverWait(driver, 30).until(expected_conditions.element_to_be_clickable((By.ID, "drivers")))
        print("[INFO]::Clicking on drivers category.")
        driversCat.click()

        element = WebDriverWait(driver, 30).until(expected_conditions.element_to_be_clickable((By.ID, "btnCollapseDriverList")))
        driverbtn = driver.find_element(By.ID, "btnCollapseDriverList")
        print("[INFO]::Expanding driver category list...")
        print("\n[WARNING]::Please be patient. This may take a minute.")
        driverbtn.click()

        element = WebDriverWait(driver, 60).until(expected_conditions.element_to_be_clickable((By.ID, "ddl-cat-btn")))
        categorybtn = driver.find_element(By.ID, "ddl-cat-btn")
        print("[INFO]::Selecting driver categories...")
        categorybtn.click()

        element = WebDriverWait(driver, 60).until(expected_conditions.presence_of_element_located((By.TAG_NAME, "tbody")))
        print("[INFO]::Located driver table.")
        element = WebDriverWait(driver, 60).until(expected_conditions.visibility_of_element_located((By.ID, "downloads-table")))
        driverlist = driver.find_element(By.ID, "downloads-table")

        driverNames = driverlist.find_elements(By.CLASS_NAME, "dl-desk-view")
        downloadlist = driverlist.find_elements(By.CLASS_NAME, "d-flex")
        driverdrp = driver.find_elements(By.NAME, "btnDriverListToggle")

        
        tableRows = driver.find_elements(By.TAG_NAME, "tr")
        tableRows = [value for value in tableRows if "tableRow" in value.get_attribute("id")]

        for i in range(len(driverNames)):
            if "BIOS" in driverNames[i].text:
                print("[INFO]::System BIOS value: {0}".format(driverNames[i].text))

                print("[INFO]::Clicking drop down...")
                driverdrp[i].click()

                driverId = tableRows[i].get_attribute("id").replace('tableRow_','child_')
                driverdescription = driver.find_element(By.ID, driverId)

                desctext = driverdescription.find_elements(By.TAG_NAME, "p")
                desctext = [desc.text for desc in desctext if desc.text != ""]

                versionsplit = desctext[0].split(',')
                version = versionsplit[0]

                execsplit = desctext[2].split('|')
                execname = execsplit[0]

                if not LABMODE:
                    result = ctypes.windll.user32.MessageBoxW(0, "Current Firmware Version: [{0}]. Upgrade to [{1}] with [{2}]?".format(CURRENTBIOSVERSION, version, execname), "Firmware Update", 4)
                    if result == 6:
                        print("[INFO]::User accepted firmware update.")
                        print("\n[INFO]::Downloading firmware updater [{0}]...".format(execname))
                        downloadlist[i].click()
                        downloadFlag = True

                        while downloadFlag:
                            files = os.listdir("C:\\Software")
                            for value in files:
                                if(value == execname):
                                    downloadFlag = False
                                    print("[INFO]::File has completed downloading.")
                            sleep(4)    

                        sleep(5)

                        return execname
                    elif result == 7:
                        print("[INFO]::User denied firmware update.")
                        sys.exit()
                    else:
                        print("[INFO]::User denied firmware update.")
                        sys.exit()
                else:
                    print("[INFO]::Lab mode is enabled, skipping user check.")
                    downloadlist[i].click()
                    downloadFlag = True

                    while downloadFlag:
                        files = os.listdir("C:\\Software")
                        for value in files:
                            if(value == execname):
                                downloadFlag = False
                                print("[INFO]::File has completed downloading.")
                        sleep(4)    

                    sleep(5)
                    return execname

    except Exception as ex:
        print("An exception occured: {0}".format(ex))
    finally:
        driver.quit()

def executeUpdater(execname):
    files = os.listdir("C:\\Software")

    hostname = socket.gethostname()
    print("[INFO]::Detecting department naming scheme...")
    if "COECIV" in hostname:
        BPPID = "REDACTED"
        print("[INFO]::Detected CIVIL naming scheme...")
    elif "COMECH" in hostname:
        BPPID = "REDACTED"
        print("[INFO]::Detected MECHANICAL naming scheme")
    elif "COELCT" in hostname:
        BPPID = "REDACTED"
        print("[INFO]::Detected ELECTRICAL naming scheme")
    elif "COITEC" in hostname:
        BPPID = "REDACTED"
        print("[INFO]::Detected INTEGRATED INFO TECH naming scheme")
    else:
        sg.theme("DefaultNoMoreNagging")


        layout = [  [sg.Text('Detected unknown hostname naming scheme. Please enter in the BIOS password to launch the update.')],
            [sg.Text('BIOS Password'), sg.InputText()],
            [sg.Button('OK', key='SUBMIT'), sg.Button('Cancel')] ]
        
        window = sg.Window('Firmware Update', layout)

        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                sys.exit()
            elif event == 'SUBMIT':
                BPPID = values[0]
                break

        window.close()


    for value in files:
        if value == execname:
            print("[INFO]::Found firmware updater: [{0}]".format(value))
            print("[INFO]::Executing firmware exec...")
            commandstring = "C:\\Software\{0} /s /f /r /p='{1}'".format(execname, BPPID)
            print("[INFO]::Executing: {0}".format(commandstring))
            if not TESTMODE:
                os.system(commandstring)
            else:
                print("[INFO]::Test mode is enabled, firmware update was not carried out.")
            break
    


firmwareUpdater = downloadBIOS()
print("[INFO]::Checking Bitlocker Status...")
if "On" in BITLOCKERSTATUS:
    print("[INFO]::Bitlocker is enabled! Stopping process...")
    ctypes.windll.user32.MessageBoxW(0, "Bitlocker needs to be suspended or turned off in order to update firmware!", "Bitlocker Enabled", 0)
    sys.exit()
else:
    print("[INFO]::Bitlocker is not enabled... continuing with process")
    executeUpdater(firmwareUpdater)
    
