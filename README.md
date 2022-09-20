# Hotkeys

Contains all files and dependencies for running NeoLab_Hotkeys.ahk, which automates importing and prepping functions.

Neolab_Hotkeys.ahk is the main script. It has dependencies on several custom .ahk libraries, organized by target program, in the \Libraries\ folder, in an attempt to keep the main script as readable as possible.

Several arrays exist in libraries next to the main script in the root folder, containing coordinates/handles/paths that could possibly be updated in future releases of target software/websites. These are seperated to make updating these easier and less risky.

Main functions include:

F1: Search for patient in 3Shape using script number from RXWizard
F2: take screenshots in 3Shape and upload them to case page on RXWizard
F3: export STLs from active case

F4: get focus on RXWizard script scan field
F5: swap between the "edit" and "review" pages
F6: retrieve order ID from myCadent site, create new note in RXWizard with It

F7: Search for patient in 3Shape using patient name and clinic name from RXWizard
F8: Create a new patient in 3Shape if none exists, create new model set otherwise
F9: Complete import process in 3Shape and return to patient browser

F10: Search for patient in myCadent site using patient name from RXWizard
F11: Download STL through OrthoCad, move and rename downloaded files to local "temp models" folder
F12: Rename files in local "temp models" folder for auto-importing, move files to queue

Functions for 3Shape model prepping:

- Select/apply plane cut, spline cut, artifact removal tools
- Select from wax knife presets
- Snap to different fixed views
- Toggle visible model
- Toggle model transparency
- Push to next step in prepping

## Installation Instructions 

Install latest version of AutoHotKey

Install Selenium 2.0.9 exe
https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
Download latest Chromedriver
replace chromedriver in SeleniumBasic install directory with new version

Create a Chrome Shortcut in Documents\Automation called ChromeForAHK with following target:
"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "--remote-debugging-port=9222"

## Update Instructions 

Locations for certain elements on the webpages or in 3Shape may change after updating them. If this happens, there are two libraries in the root directory that contain almost all of these locations, Web_Paths and Ortho_Locations. Open them and replace the necessary variables inside their arrays


