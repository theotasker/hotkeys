# Neolab Importing and Prepping Assistance Hotkeys
AutoHotKey(AHK)/Selenium program to assist in the importing and model prep workflows at Neolab by automating web browsing, data entry, and file management. Should not be run outside the Neolab environment. Interacts with the following software/websites:
- RXWizard: Neolab's proprietary web portal/lab management website
- myCadent: web portal for retreiving intraoral scans (STLs) from the iTero scanners
- OrthoCad: Standalone exe used to securely download/export STLs from myCadent
- 3Shape OrthoAnalyzer/ApplianceDesigner: Dental/Orthodontic specific CAD software for preparing STLs, designing appliances, etc.
- 3shape OrthoSystem: New updated version with a new interface that allows for command line interactions

## Developer/Installer Information
Neolab_Hotkeys.ahk is the main script. It has been reduced down as much as possible to improve readability. Most of its contents are simply assigning navigation/data entry functions to be triggered by the F1-F12 keys, assigning 3Shape CAD functions to be triggered through 3D mouse buttons, and managing the AHK progress bar.

It has dependencies on several custom .ahk libraries, organized by target software, in the \Libraries\ folder, in an attempt to keep the main script as readable as possible. Also contained in the \Libraries\ folder is the CaptureScreen.ahk library, largely copied from the Gdip standard library v1.45 by tic (Tariq Porter). Note that AutoHotKey must be run in 32-bit mode for this library to function.

Several libraries sit next to the main script in the root folder, containing coordinates/handles/paths that could possibly be updated in future releases of target software/websites. These are seperated to make updating them easier and less risky.

#### Purpose Of 3D Mouse Functions
3Shape OrthoAnalyzer/ApplianceDesigner can be quite frustrating as a CAD operator. It is compatible with a 3D mouse, but there are almost no default key bindings, and it lacks the ability to set key bindings manually. And the model prepping process requires *many* function activations through mouse clicks. The 3shape dev team has stated they're not making any significant updates to the ortho software, so the next best option was to replicate mouse clicks with the AutoHotKey software. 

### Installation Instructions 
- Install latest version of AutoHotKey in 32-bit mode
- Install Selenium 2.0.9 exe, https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
- Download latest Chromedriver, replace chromedriver in SeleniumBasic install directory with new version
- Install latest Chrome
- Create a Chrome Shortcut in %user%\Documents\Automation\ called ChromeForAHK with following target: "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "--remote-debugging-port=9222"
- Place a shortcut to Neolab_Hotkeys.ahk in the user's startup folder, so the program runs in the background

### Update Instructions 
Locations for certain elements on the webpages or in 3Shape may change after updates. This may result in an error message if AHK is unable to find the element, or unexpected behavior if it points to the wrong place. If this happens, there are two libraries in the root directory that contain all of these locations, Web_Paths and Ortho_Locations. Open them and replace the necessary variables inside their arrays.

## User Information
This software is intended for use by the importing and prepping teams. It is compatible with both the 2019 and 2021 version of 3Shape, and with both the webtech and CSR roles on RXWizard.

### WebDriver Basics
In order to use the data entry/navigation functions, the user must first launch a webdriver, which is an instance of Chrome controlled by the script. This requires a very specific configuration, so there's not much flexibility in deployment. Each instance of a webdriver must be in its own window, with no other tabs residing beside it. Neolab_Hotkeys is capable of controlling two webdrivers, one for RXWizard, and one for myCadent. It's recommended that the user use Firefox for web browsing unrelated to these sites.

To open a new webdriver, simply press an F key that relies on the respective website. For example, pressing F4 for the first time will prompt you if you want to open a new webdriver for RXWizard, and F10 will do so for myCadent. If you have other Chrome instances open, it will prompt you if you want to close them. This is recommended. 

### Main functions
- F1: Search for patient in 3Shape using script number from RXWizard
- F2: take screenshots in 3Shape and upload them to case page on RXWizard
- F3: export STLs from active case

- F4: get focus on RXWizard script scan field
- F5: swap between the "edit" and "review" pages
- F6: retrieve order ID from myCadent site, create new note in RXWizard with It

- F7: Search for patient in 3Shape using patient name and clinic name from RXWizard
- F8: Create a new patient in 3Shape if none exists, create new model set otherwise
- F9: Complete import process in 3Shape and return to patient browser

- F10: Search for patient in myCadent site using patient name from RXWizard
- F11: Download STL through OrthoCad, move and rename downloaded files to local "temp models" folder
- F12: Rename files in local "temp models" folder for auto-importing, move files to queue

### 3D Mouse Functions Inside 3Shape Model Prep
- Select/apply plane cut, spline cut, artifact removal tools
- Select from wax knife presets
- Snap to different fixed views
- Toggle visible model
- Toggle model transparency
- Push to next step in prepping






