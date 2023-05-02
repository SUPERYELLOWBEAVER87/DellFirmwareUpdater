## Introduction :wave:
This program fully automates updating BIOS Firmware on Dell specific machines through the usage of Selenium. The program works by using selenium to download the latest firmware update from Dell's warranty and contracts, then launching the program through CMD. The practicality of this is mostly so that system administrators and IT managers can fully automate updating firmware in large computer labs that have different versions and models of Dell machines.

## Requirements
The program is packaged and compiled with PyInstaller into a portable standalone executable. Chromium is used as the web driver of choice for this project.
### Development Environment
* Selenium
* webdriver_manager
* PySimpleGUI

![image](https://user-images.githubusercontent.com/78384615/235709838-10d02099-b329-4f6e-b1f3-a6d5bfda41c3.png)

## Implementation
The general code workflow is as follows: First the program makes sure all pre-requisites are satisfied, Chrome to be installed, as well as administrator permissions in order to check the status of Bitlocker. Once the program confirms that all pre-requisites are satisfied, it proceeds by launching a chrome webdriver sending the local serial through the lookup, then downloads the BIOS firmware for the specific machine. Afterwards, it ensures that Bitlocker is disabled before proceeding with launching the BIOS firmware update executable with appropriate handles to execute silently.
