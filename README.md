# Neolab Importing and Prepping Assistance Hotkeys
AutoHotKey(AHK)/Selenium program to assist in the importing and model prep workflows at Neolab by automating web browsing, data entry, and file management. Should not be run outside the Neolab environment. Interacts with the following software/websites:
- RXWizard: Neolab's proprietary web portal/lab management website
- myCadent: web portal for retreiving intraoral scans (STLs) from the iTero scanners
- OrthoCad: Standalone exe used to securely download/export STLs from myCadent
- 3Shape OrthoAnalyzer/ApplianceDesigner: Dental/Orthodontic specific CAD software for preparing STLs, designing appliances, etc.
- 3shape OrthoSystem: New updated version with a new interface that allows for command line interactions

## Developer/Installer Information
Neolab_Hotkeys.ahk is the main script. It has been reduced down as much as possible to improve readability. Most of its contents are simply assigning navigation/data entry functions to be triggered by the F1-F12 keys, assigning 3Shape CAD functions to be triggered through 3D mouse buttons, and managing the AHK progress bar.

It has dependencies on several custom .ahk libraries, organized by target software, in the \Libraries\ folder. Also contained in the \Libraries\ folder is the CaptureScreen.ahk library, largely copied from the Gdip standard library v1.45 by tic (Tariq Porter). Note that AutoHotKey must be run in 32-bit mode for this library to function.

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

### Progress Bar
Functions that take more than 1/10 of a second are accompanied by a progress bar. Most of these functions will turn off the user's control over the mouse for the duration, to prevent the user from interrupting any active automation process. AHK can not, however, interrupt keystrokes in this way, so it is imperative that the user refrain from doing anything while a progress bar is visible on the screen.

If one of these longer functions fails, the user will be left without mouse control for at most 30 seconds, while the program tries to find the respective element. After this, a message box will tell the user what failed, and mouse control will be restored.

### WebDriver Basics
In order to use the data entry/navigation functions, the user must first launch a webdriver, which is an instance of Chrome controlled by the script. This requires a very specific configuration, so there's not much flexibility in deployment. Each instance of a webdriver must be in its own window, with no other tabs residing beside it. Neolab_Hotkeys is capable of controlling two webdrivers, one for RXWizard, and one for myCadent. It's recommended that the user use Firefox for web browsing unrelated to these sites.

To open a new webdriver, simply press an F key that relies on the respective website. For example, pressing F4 for the first time will prompt you if you want to open a new webdriver for RXWizard, and F10 will do so for myCadent. If you have other Chrome instances open, it will prompt you if you want to close them. This is recommended. 

### Main functions
These functions do not require any specific software to be active. AHK will get focus on any software it needs to interact with, as long as it exists, and does not have a pop-up window preventing it from progressing.

- F1: Retrieves patient info from current RXWizard case page, searches for patient in 3Shape by script number
- F2: Takes screenshots of currently open case in 3Shape from the front and both sides, and uploads them to the currently open RXWizard case
- F3: (no use currently)
- F4: Press at any time to push focus to the scan field on RXWizard, then scan a script barcode to navigate to that case
- F5: While on a case page on RXWizard, swaps between the "edit" and "review" pages
- F6: Retrieves order ID from myCadent patient scan page, creates a new note with it on the current RXWizard case
- F7: Retrieves patient info from current RXWizard case page, searches for patient in 3Shape by patient name and clinic name
- F8: Retrieves patient info from current RXWizard case page, creates new patient in 3Shape if patient doesn't exist, otherwise creates model set
- F9: After importing STLs into 3Shape, returns user to patient browser, and deletes STLs in user's "temp models" folder
- F10: Retrieves patient info from current RXWizard case page, searches myCadent for scan using patient's name, last comma first
- F11: While on patient scan page on myCadent, clicks to downloads, finalizes OrthoCad export, and moves STLs to user's "temp models" folder
- F12: Retrieves patient info from current RXWizard case page, and prompts user on which arches should be finished, and whether auto-importing should be used

### 3D Mouse Functions Inside 3Shape Model Prep
- "Fit" button: Presses "next" button during model prep. For last step, also initiates STL export from patient browser.
- "T", "F", "R" buttons: Clicks snap view buttons. Single press clicks default view, double press clicks alternate view, i.e. "front" vs "back"
- "Rotate Lock" button: Toggles model transparency
- "Roll" button: Toggles visible models
- "CTRL" button: Activates the "plane cut" tool if it's not active, applies the tool if it is active
- "ALT" button: Activates the "spline cut" tool if it's not active and sets it to "smooth", applies the tool if it is active
- "SHIFT" button: Activates the "artifact removal" tool if it's not active, applies the tool if it is active
- "1"-"4" buttons: Activates the respective "wax knife" preset. Double tap to select the alternate preset, i.e. double-tapping 1 selects preset 5.