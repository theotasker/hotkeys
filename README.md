# Hotkeys

Contains all scripts and libraries needed for running the "Neolab Hotkeys" AutoHotKey program. 

Neolab_Hotkeys.ahk is the main script, it has dependencies on several .ahk libraries in the \Libraries\ folder. Users should be given the shortcut in the top level folder, so it will run the latest version 

Main functions of the script include automating data entry/retreival between RXWizard, MyCadent, and 3Shape for importers, enabling the use of the 3D mouse buttons for functions in 3Shape (by replicating mouse clicks), prepping files for auto importing, and taking occlusion screenshots



## Installation Instructions 

Selenium 2.0.9 exe

https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0

Update chromedriver in install directory with new version

Chrome Shortcut update

/Properties/Target:   "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "--remote-debugging-port=9222"
