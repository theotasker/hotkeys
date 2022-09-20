; ===========================================================================================================================
; Library for Cadent_ functions
; ===========================================================================================================================

global Cadentcheck := 0
global MyCadentDriver := ""

; ===========================================================================================================================
; myCadent site functions
; ===========================================================================================================================

Cadent_StartWebDriver()
{
	if !FileExist(chromeShortcutDir "ChromeForAHK.lnk")
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Chrome Shortcut Not Found, Can't find the proper shortcut at %A_MyDocuments%\Automation\ChromeForAHK.lnk `nRemember to modify the shortcut to run in debug mode by adding "--remote-debugging-port=9222"
		Exit
	}

	Run, ChromeForAHK.lnk, %chromeShortcutDir%
	Sleep, 500
	global MyCadentDriver := ChromeGet()
	MyCadentDriver.SwitchToWindowByTitle("New Tab")
	MyCadentDriver.Get("https://mycadent.com/COrdersList.aspx")

	CadentCheck := 1

	return MyCadentDriver
}

Cadent_StillOpen() ; Checks to see if cadent driver is open, reopens if not
{
	if CadentCheck = 0
	{
		MsgBox, 4, Open CadentDriver?, No instance of CadentDriver found, initiate?
			IfMsgBox, Yes
			{
				global MyCadentDriver := Cadent_StartWebDriver()
				return False
			}
			IfMsgBox, No
			{
				Exit
			}
	}

	try currentURL := MyCadentDriver.Url
	catch e
	{
		MsgBox, 4, Webdriver Error, The tab for driving MyCadent was closed, reinitiate webdriver?
			IfMsgBox, Yes
				global MyCadentDriver := Cadent_StartWebDriver()
				return False
			IfMsgBox, No
			{
				return
			}
		Exit
	}
	return currentURL
}

Cadent_ordersPage(patientInfo, patientSearch) ; goes to the patient search page, enters patient info if asked
{
	if (MyCadentDriver.Url != "https://mycadent.com/COrdersList.aspx")
    {
        MyCadentDriver.Get("https://mycadent.com/COrdersList.aspx")
    }

	BlockInput MouseMove

	if WinExist("Clinical Orders List - Google Chrome")
		WinActivate
	else
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox The MyCadent site must be on the top tab for this shortcut to work
		Exit
	}

	Send, {tab}{tab}{tab}
	sleep, 100

    MyCadentDriver.findElementByID(cadentCssID["searchField"]).click()
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}
	sleep, 100

	if (patientSearch = True) {
		if StrLen(patientInfo["firstName"]) > 1 and StrLen(patientInfo["lastName"]) > 1
		{
			Send % patientInfo["lastName"]
			Send {,}{space}
			Send % patientInfo["firstName"]
		}
	}
	BlockInput, MouseMoveOff
	return
}

Cadent_GetOrderID()
{
	Cadent_StillOpen()

	if !InStr(MyCadentDriver.Url, "https://mycadent.com/CaseInfo.aspx")
	{
		BlockInput MouseMoveOff
		Gui, Destroy		
		MsgBox Must be on MyCadent order page to use this function
		Exit
	}

	orderID := MyCadentDriver.findElementByID(cadentCssID["orderID"]).Attribute("value")

	return orderID
}

Cadent_exportClick(currentURL) 
{
	BlockInput MouseMove

	if !InStr(currentURL, "https://mycadent.com/CaseInfo.aspx") {
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Wrong Page, Must be on a case page in MyCadent
		Exit
	}

	try MyCadentDriver.findElementByID(cadentCssID["exportLink"]).click() 
	catch e {
		BlockInput MouseMoveOff
		Gui, Destroy		
		Msgbox,, Web Error, Couldn't click on the export button on MyCadent
		Exit
	}

	BlockInput, MouseMoveOff
	return
}

; ===========================================================================================================================
; OrthoCad functions
; ===========================================================================================================================

Cadent_exportOrthoCAD(patientInfo) 
{
	BlockInput, MouseMove

	SetTitleMatchMode, 1
	WinWait, ahk_exe OrthoCAD.exe,, 30, Export  ; Wait for main window to open
	if ErrorLevel {
			BlockInput MouseMoveOff
			Gui, Destroy
			MsgBox,, OrthoCAD Error, Main OrthoCAD window didn't open
			Exit
		}

	Sleep, 400 ; buffer for file to load
	Cadent_ordersPage(patientInfo, patientSearch:=False)
	BlockInput, MouseMove

	SetTitleMatchMode, 3
	WinWait, OrthoCAD Export,, 30   ; Wait for export box to open
	if ErrorLevel {
			BlockInput MouseMoveOff
			Gui, Destroy
			MsgBox,, OrthoCAD Error, Export box didn't open
			Exit
		}

    WinActivate, ahk_exe OrthoCAD.exe
    WinActivate, "OrthoCAD Export"

	sleep, 500

    ControlFocus, ComboBox1, OrthoCAD Export   ; set the export type to open scan
    Send, {down}
    Sleep, 100

    ControlFocus, ComboBox2, OrthoCAD Export ; set the export type to models in occlusion
    Send, {down}{down}
    Sleep, 100

    ControlFocus, Edit1, OrthoCAD Export ; set the export folder
    Send {CtrlDown}a{CtrlUp}
    Sleep, 100

	exportFilename := patientInfo["firstName"] patientInfo["lastName"]
    Send, % exportFilename
    Sleep, 100

    Send, {tab} ; go to the export button and hit enter
    Sleep, 100
    Send, {enter}

	SetTitleMatchMode, 3
	WinWait, Export Done,, 30
	if ErrorLevel {
			BlockInput MouseMoveOff
			Gui, Destroy
			MsgBox,, OrthoCAD Error, Export confirmation box didn't open
			Exit
		}

	WinActivate, Export Done
	WinWaitActive, Export Done,, 30
	if ErrorLevel {
			BlockInput MouseMoveOff
			Gui, Destroy
			MsgBox,, OrthoCAD Error, Couldn't get focus on export confirmation box
			Exit
		}

    Sleep, 100
	ControlFocus, Button1, Export Done  ; decline to navigate to file location
	Sleep, 300
	Send, {tab}
	Sleep, 100
	Send, {enter}
	Sleep, 100

	WinKill, ahk_exe OrthoCAD.exe
	Sleep, 200

	SetTitleMatchMode, 2
	if WinExist("ahk_class #32770")  ; confirm closing without saving
	{  
		Send {tab}
		sleep, 100
		Send {enter}
		Sleep, 200
	}

	WinWaitClose, ahk_exe OrthoCAD.exe,, 10
	if ErrorLevel 
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, OrthoCAD Error, Couldn't close OrthoCAD
		Exit
	}

	BlockInput, MouseMoveOff
	return exportFilename
}

Cadent_moveSTLs(exportFilename) 
{
	IfNotExist, C:\Cadent\Export\ 
	{
		FileCreateDir, C:\Cadent\Export\
	}

	IfNotExist, C:\Cadent\Export\%exportFilename% 
	{
		BlockInput MouseMoveOff
		Gui, Destroy		
		MsgBox Couldn't find export folder, didn't move files
	}

	if FileExist("C:\Cadent\Export\" exportFilename "\" "*u.stl") 
	{
		FileMove, C:\Cadent\Export\%exportFilename%\*u.stl, %tempModelsDir%%exportFilename%upper.stl, 1
		counter += 1
	}

	if FileExist("C:\Cadent\Export\" exportFilename "\" "*l.stl") 
	{
		FileMove, C:\Cadent\Export\%exportFilename%\*l.stl, %tempModelsDir%%exportFilename%lower.stl, 1
		counter += 1
	}

    if FileExist("C:\Cadent\Export\" exportFilename "\" "*.stl") 
	{
		FileMove, C:\Cadent\Export\%exportFilename%\*.stl, %tempModelsDir%, 1
		counter += 1
	}

    IfExist, C:\Cadent\Export\%exportFilename% 
	{
        FileRemoveDir, C:\Cadent\Export\%exportFilename%, 1
	}

	return
}