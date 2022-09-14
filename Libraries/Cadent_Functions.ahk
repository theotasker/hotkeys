Cadentcheck := 0

Cadent_StartWebDriver()
{
	global

	; Check to make sure that the modified Chrome shortcut is in the right place
	chromelocation = %A_MyDocuments%\Automation\ChromeForAHK.lnk
	if !FileExist(chromelocation)
	{
		MsgBox,, Chrome Shortcut Not Found, Can't find the proper shortcut at %A_MyDocuments%\Automation\ChromeForAHK.lnk `nRemember to modify the shortcut to run in debug mode by adding "--remote-debugging-port=9222"
		Exit
	}

	; Open another window and assign it to mycadent
	Run, ChromeForAHK.lnk, %A_MyDocuments%\Automation\
	Sleep, 500
	MyCadentDriver := ChromeGet()
	MyCadentDriver.Get("https://mycadent.com/COrdersList.aspx")

	CadentCheck := 1

	Gui, Destroy
	Exit
}

; function to check that Cadent is still open
Cadent_StillOpen()
{
	global MyCadentDriver, CadentCheck

	if CadentCheck = 0
	{
		Gui, Destroy
		MsgBox, 4, Open CadentDriver?, No instance of CadentDriver found, initiate?
			IfMsgBox, Yes
			{
				Cadent_StartWebDriver()
			}
			IfMsgBox, No
			{
				Gui, Destroy
				Exit
			}
	}

	try currentURL := MyCadentDriver.Url
	catch e
	{
		MsgBox, 4, Webdriver Error, The tab for driving MyCadent was closed, reinitiate webdriver?
			IfMsgBox, Yes
				Cadent_StartWebDriver()
			IfMsgBox, No
			{
				return
			}
		Gui, Destroy
		Exit
	}
	return currentURL
}

; function to go to orders page
Cadent_ordersPage(patientInfo, patientSearch)
{
	global MyCadentDriver, Path_CadentSearchFieldID
	Cadent_StillOpen()

	; Check the current page of the cadent driver
	if (MyCadentDriver.Url != "https://mycadent.com/COrdersList.aspx")
    {
        MyCadentDriver.Get("https://mycadent.com/COrdersList.aspx")
    }

	if WinExist("Clinical Orders List - Google Chrome")
		WinActivate
	else
	{
		MsgBox The MyCadent site must be on the top tab for this shortcut to work
		Gui, Destroy
		Exit
	}

	Sleep, 100
	Send, {tab}
	Sleep, 100
    MyCadentDriver.findElementByID(Path_CadentSearchFieldID).click()
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}

	if (patientSearch = True) {
		if StrLen(patientInfo["firstName"]) > 1 and StrLen(patientInfo["lastName"]) > 1
		{
			BlockInput, MouseMove
			Send % patientInfo["lastName"]
			Send ,{space}
			Send % patientInfo["firstName"]
			BlockInput, MouseMoveOff
		}
	}
	return
}

Cadent_exportClick() {
	global MyCadentDriver, Path_CadentExport
	if !InStr(currentURL, "https://mycadent.com/CaseInfo.aspx") {
		MsgBox,, Wrong Page, Must be on a case page in MyCadent
		Exit
	}

	try MyCadentDriver.findElementByID(Path_CadentExport).click() ; click on the export button
	catch e {
		Msgbox,, Web Error, Couldn't click on the export button on MyCadent
		Exit
	}
	return
}

Cadent_exportOrthoCAD(patientInfo) {
	BlockInput, MouseMove

	SetTitleMatchMode, 1
	WinWait, ahk_exe OrthoCAD.exe,, 30, Export  ; Wait for main window to open
	if ErrorLevel {
			BlockInput, MouseMoveOff
			MsgBox,, OrthoCAD Error, Main OrthoCAD window didn't open
			Exit
		}

	Sleep, 400 ; buffer for file to load
	Cadent_ordersPage(patientInfo:="", patientSearch:=False)

	SetTitleMatchMode, 3
	WinWait, OrthoCAD Export,, 30   ; Wait for export box to open
	if ErrorLevel {
			BlockInput, MouseMoveOff
			MsgBox,, OrthoCAD Error, Export box didn't open
			Exit
		}

    WinActivate, ahk_exe OrthoCAD.exe
    WinActivate, "OrthoCAD Export"

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
			BlockInput, MouseMoveOff
			MsgBox,, OrthoCAD Error, Export confirmation box didn't open
			Exit
		}

	WinActivate, Export Done
	WinWaitActive, Export Done,, 30
	if ErrorLevel {
			BlockInput, MouseMoveOff
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
	if WinExist("ahk_class #32770") {  ; confirm closing without saving
		Send {tab}
		sleep, 100
		Send {enter}
		Sleep, 200
	}

	WinWaitClose, ahk_exe OrthoCAD.exe,, 10
	if ErrorLevel {
		BlockInput, MouseMoveOff
		MsgBox,, OrthoCAD Error, Couldn't close OrthoCAD
		Exit
	}
	return exportFilename
}

Cadent_moveSTLs(exportFilename) {
	IfNotExist, C:\Cadent\Export\ {
		FileCreateDir, C:\Cadent\Export\
	}

	IfNotExist, %A_MyDocuments%\Temp Models {
		FileCreateDir, %A_MyDocuments%\Temp Models
	}

	IfNotExist, C:\Cadent\Export\%exportFilename% {
		MsgBox Couldn't find export folder, didn't move files
	}

	if FileExist("C:\Cadent\Export\" exportFilename "\" "*u.stl") {
		FileMove, C:\Cadent\Export\%exportFilename%\*u.stl, %A_MyDocuments%\Temp Models\%exportFilename%upper.stl, 1
		counter += 1
	}

	if FileExist("C:\Cadent\Export\" exportFilename "\" "*l.stl") {
		FileMove, C:\Cadent\Export\%exportFilename%\*l.stl, %A_MyDocuments%\Temp Models\%exportFilename%lower.stl, 1
		counter += 1
	}

    if FileExist("C:\Cadent\Export\" exportFilename "\" "*.stl") {
		FileMove, C:\Cadent\Export\%exportFilename%\*.stl, %A_MyDocuments%\Temp Models\, 1
		counter += 1
	}

    IfExist, C:\Cadent\Export\%exportFilename% {
        FileRemoveDir, C:\Cadent\Export\%exportFilename%, 1
	}

	return
}

Cadent_GetOrderID()
{
	global MyCadentDriver, Path_CadentOrderNumberID

	Cadent_StillOpen()

	if !InStr(MyCadentDriver.Url, "https://mycadent.com/CaseInfo.aspx")
	{
		MsgBox Must be on MyCadent order page to use this function
		Gui, Destroy
		Exit
	}

	orderID := MyCadentDriver.findElementByID(Path_CadentOrderNumberID).Attribute("value")

	return orderID

}