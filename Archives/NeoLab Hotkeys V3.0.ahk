; ===========================================================================================================================
; Combined functions for Importing, Prepping, and Engraving, using netfabb, 3Shape, and the websites
; ===========================================================================================================================

; ================================================================================================================
; Selenium Paths, need updating as RXWizard gets updated
; ================================================================================================================

; ---------------------------------------------------------------------------------------------------------------------
; RXWizard Paths

Path_GenerateScheduleCSS = #main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button:nth-child(1)

Path_StartStopXPATH = //*[@id="main-content-wrapper"]/div[1]/main/div[2]/div/div[1]/button/span

Path_StartStopCSS = #main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button

; If user has manager privileges on the website, the button location changes
Path_StartStopManagerXPATH = //*[@id="main-content-wrapper"]/div[1]/main/div[2]/div/div/div[2]/button[1]

Path_SaveNoteXPATH = /html/body/div[4]/div/div[2]/div/div[2]/div[2]/div/button[2]

Path_UploadFileXPATH = //*[@id="main-content-wrapper"]/div[1]/main/div[1]/div[1]/form/div/div/div/div[2]/span/div[1]/span/div/p[2]

Path_ScanScriptCSS = #react > section > header > form > div > input

Path_ReviewButtonCSS = #main-content-wrapper > div > main > form > div.ant-row-flex.ant-row-flex-end > div:nth-child(2) > button

Path_ScriptNumberCSS = #main-content-wrapper > div.page-content > div > div.name > div > span.case-name

Path_ClinicNameCSS = #main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(1) > div.ant-col.ant-col-17 > div > div

Path_PatientNameCSS = #main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(1) > div.card > div > div:nth-child(8) > div.ant-col.ant-col-17 > div > div

Path_NewNoteCSS = #main-content-wrapper > div.page-content > main > div.ant-row > div:nth-child(2) > div:nth-child(4) > div > button

; ---------------------------------------------------------------------------------------------------------------------
; Mycadent paths

Path_CadentSearchFieldID = ctl00_body_OrdersListReport_ctl01_ctl05_ctl05-string-operand

Path_CadentOrderNumberID = ctl00_body_txtOrderHeaderID

Path_CadentExport = ctl00_body_ucOrthoCadLink_OrthoCadLink

; ===========================================================================================================================
; AutoHotKey Startup Stuff
; ===========================================================================================================================


#SingleInstance, Force

; Library specifically to capture screenshots, must be in 32bit mode in AHK
#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\CaptureScreen.ahk

; starter variable for the F2 function in netfabb
placedengraving =: 0

; Sets the relative click and pixel points to relative to the client window, seems to be most repeatable
CoordMode, Mouse, Client
CoordMode, Pixel, Client

; Variables set at the start to see if any webdrivers have been started previously
NeoCheck := 0
Cadentcheck := 0

; Defines the location for a txt which will determine the start state of the "manager" check box, creates one if none
ManagerCheckFile := A_MyDocuments "\Automation\manager.txt"
if !FileExist(ManagerCheckFile)
	FileAppend, no, %ManagerCheckFile%

FileRead, ManagerCheckText, %ManagerCheckFile% ; get the value from the txt file to determine start state of check box
	
; Sets the start state of the manager check
if ManagerCheckText = no
{
	managercheck := "no"
}
else if ManagerCheckText = yes
{
	managercheck := "yes"
}
else 
{
	MsgBox,, File Error, Error with the Manager Check text in Documents\Automation
	Exit
}

Log_ImportLocalLog() ; Imports the current log count from the local log, creates a log if none exists

SetBox() ; opens the settings box to select step

; =========================================================================================================================
; Other starter functions
; =========================================================================================================================

; Creates group to select either ortho or appliance designer
GroupAdd, ThreeShape, OrthoAnalyzer
GroupAdd, ThreeShape, ApplianceDesigner

GroupAdd, ThreeShapePatient, New patient model set info
GroupAdd, ThreeShapePatient, New patient info

GroupAdd, ThreeShapeExe, ahk_exe OrthoAnalyzer.exe
GroupAdd, ThreeShapeExe, ahk_exe ApplianceDesigner.exe

; Variable for the prepping steps in orthoan/appliance designer, used for the next step clicks
3ShapeSteps := "Prepare occlusion,Setup plane alignment,Virtual base,Sculpt maxillary,Sculpt mandibular"

; sets start variables for use in the double button tap functions
Var1 := A_TickCount
Var2 := A_TickCount
Var3 := A_TickCount
VarT := A_TickCount
VarR := A_TickCount
VarF := A_TickCount

; Directories
Import_Input_Dir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Input\"
ABR_Input_Dir := "\\app03\Scans\~Digital Dept Share\R&D Network\EasyRX\Auto Import Input\"

; =========================================================================================================================
; Progress bar utilities and step settings
; =========================================================================================================================

; default progress bar location
ProgX := "x788"
ProgY := "y150"

; default step for the program
Step := "Prepping"
StartStep := "Start Digital Prep"
EndStep := "End Digital Prep"

Pause::
{
	SetBox()
	return
}

Progress: ; Set the location of the progress bar
{
	ProgX := "x788"
	ProgY := "y150"
	Gui, Destroy
	Gui, Add, Text, x0 y0 w320 h15 , Move this box to where you want your progress bar to be
	Gui, Add, Button, x52 y15 w120 h20 gSetLocation, Set Location
	Gui, Show, %ProgX% %ProgY% w320 h30, Select Location
	return
}

SetLocation: ; Subroutine for the progress bar location set GUI
{
	Gui, Show
	WinGetPos, VarX, VarY,,, Select Location
	ProgX := "x" + VarX
	ProgY := "y" + VarY
	Gui, Destroy
	return
}

Importing:
{
	Step := "Importing"
	StartStep := "Start Digital Import"
	EndStep := "End Digital Import"
	Gui, Destroy
	return
}

Prepping:
{
	Step := "Prepping"
	StartStep := "Start Digital Prep"
	EndStep := "End Digital Prep"
	Gui, Destroy
	Exit
}

Engraving:
{
	Step := "Engraving"
	StartStep := "Start Engraving"
	EndStep := "End Engraving"
	Gui, Destroy
	Exit
}

Printing:
{
	Step := "Printing"
	StartStep := "Start 3D Print Model"
	EndStep := "End 3D Print Model"
	Gui, Destroy
	Exit
}

GuiClose:
{
	Gui, Destroy
	Exit
}

ManagerCheckSub: ; runs when the check box is clicked, updates txt file and updates program variable for website
{
	Gui, Submit, NoHide
	
	if ManagerCheckVar = 1
	{
		FileDelete, %ManagerCheckFile%
		FileAppend, yes, %ManagerCheckFile%
		managercheck := "yes"
	}
	if ManagerCheckVar = 0
	{
		FileDelete, %ManagerCheckFile%
		FileAppend, no, %ManagerCheckFile%
		managercheck := "no"
	}
	return
}

Guide:
{
	Run, https://docs.google.com/document/d/1nQaEP69PKuJdn2lcOLDvR4fMiKJfQNFgxwkA2em04HM/edit?usp=sharing
	return
}


SetBox()
{
	global	
	Gui, Add, Button, x12 y210 w100 h30 gImporting, Importing
	Gui, Add, Button, x112 y210 w100 h30 gPrepping, Prepping
	Gui, Add, Button, x212 y210 w100 h30 gEngraving, Engraving
	Gui, Add, Button, x312 y210 w100 h30 gPrinting, Printing
	Gui, Add, Button, x82 y10 w250 h30 gProgress, Change Progress Bar Location
	Gui, Add, Text, x82 y180 w250 h20 +Center, Change Current Step
	Gui, Add, Text, x5 y80 w250 h100, Cases imported = %importedlocal%`nCases prepped = %preppedlocal%`nCases engraved = %engravedlocal%`nHotkeys used = %hotkeyslocal%`nMouse actions saved = %clickslocal%`nKeystrokes saved = %strokeslocal%
	
	FileRead, ManagerCheckText, %ManagerCheckFile% ; get the value from the txt file to determine start state of check box
	
	if ManagerCheckText = no
	{
		Gui, Add, Checkbox, x130 y50 w250 h20 vManagerCheckVar gManagerCheckSub, "Manager" account on rxwizard?
	}
	else if ManagerCheckText = yes
	{
		Gui, Add, Checkbox, x130 y50 w250 h20 vManagerCheckVar gManagerCheckSub Checked, "Manager" account on rxwizard?
	}
	else 
	{
		MsgBox,, File Error, Error with the Manager Check text in Documents\Automation
		Exit
	}
	
	Gui, Font, underline
	Gui, Add, Text, cBlue gGuide x162 y120 w120 h14 +Center, Click here to view guide
	Gui, Show, w439 h253, NeoLab AutoHotKey Settings
	return
}

; =========================================================================================================================
; Logging shortcuts
; =========================================================================================================================


^k::
{
   Log_Increment(6, 1)
   Log_UpdateServer()
   Log_UpdateLocal()
   return
}



; ===========================================================================================================================
; Engraving shortcuts in Netfabb
; ===========================================================================================================================

; starter variable for the F2 function in netfabb
placedengraving =: 0

#IfWinNotActive ahk_exe explorer.exe

f1::
{
	Netfabb_Level()
	Log_Increment(6, 4)
	return
}

f2::
{
	Netfabb_finish()
	Log_Increment(4, 0)
	return
}


#IfWinNotActive ahk_class Notepad


; =========================================================================================================================
; Auto Importing finishing
; =========================================================================================================================



; =========================================================================================================================
; RXWizard Shortcuts
; =========================================================================================================================

; shortcut to call the home function and return to the cases page
f4::
{
    Neo_NavigateToCases()
	Log_Increment(2, 0)
    return
}

; shortcut to navigate to review from edit page, must be on edit page
f5::
{
    ; Neo_NavigateReviewFromEdit()
	; Log_Increment(2, 0)
	; Gui, Destroy
    ; return

	temppath = #main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button > span

	temppath2 = #main-content-wrapper > div.page-content > main > div.card > div > div:nth-child(1) > button

	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.CSS_SELECTOR, temppath2))
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the start button before it was clicked the first time
		Exit
	}

	
	
	webstep := NeoDriver.findElementByCss(temppath2).Attribute("innerText")

	NeoDriver.findElementByCss(temppath2).sendKeys(Keys.RETURN)

	
	return
}

; shortcut to hit start stop on a page, must be on review page. 
f6::
{
    global NeoDriver, currentstep, currentstepcss, currentstepxpath, ProgX, ProgY, StartStep, EndStep, Path_StartStopXPATH, Path_StartStopManagerXPATH, Path_GenerateScheduleCSS
    
    Gui, Add, Progress, vprogress w300 h45
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop
    
    if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/") ; must be on review page, or recieve error
    {      
		Gui, Destroy
        MsgBox,, Wrong Page, Must be on the review page for this function
        Exit
    }        
		
		
	; if managercheck = no ; checks for the manager.txt doc, will click the generate schedule button if not
	; generate schedule seems to have gone away on the webtech role, commented this out
		
		; Click the "generate schedule" button
		; try NeoDriver.findElementByCss(Path_GenerateScheduleCSS).Click()
		; catch e 
		; {
			; MsgBox,, Couldn't Find Element, Couldn't find the "Generate Schedule" button
			; Gui, Destroy
			; Exit
		; }
	
        
	GuiControl,, Progress, 33
    
	; Current step variables
	currentstep := StartStep
	
	
	if managercheck = no
	{
		; xpath for non-managers
		currentstepxpath := Path_StartStopCSS
	}
	if managercheck = yes
	{
		; xpath for managers
		currentstepxpath := Path_StartStopManagerXPATH
	}
	
	try NeoDriver.executeScript("window.scrollBy(0, 800)")
	catch e
	{
		msgbox couldn't scroll
		return
	}
    
	; calls Neo_ function to click the start button
	Neo_Start2()
        
	GuiControl,, Progress, 66
    
	currentstep := EndStep
    
	; calls Neo_ function to click the stop button
	Neo_Stop2()
	
	Sleep, 1000
        
	GuiControl,, Progress, 100
	Gui, Destroy
		
	Neo_NavigateToCases()
	
	Log_Increment(5, 0)
	
	Log_UpdateServer()
	Log_UpdateLocal()
    
	return
}

; =========================================================================================================================
; Ortho Analyzer Shortcuts
; =========================================================================================================================


; shortcut to paste and search inside the advanced search dialogue
f7::
{
	Neo_GetInfoFromReview()
    BlockInput, MouseMove
	
	if Step = Importing
	{
		Ortho_AdvSearch()
		Log_Increment(6, 15)
	}
	else
	{
		Ortho_OpenCase()
		Log_Increment(6, 1)
	}
	

	
	BlockInput, MouseMoveOff
		
    return
}

; shortcut to paste patient and clinic into new patient info fields
f8::
{
    global scriptnumber, ProgX, ProgY
    
    Gui, Add, Progress, vprogress w300 h45
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop
    
    if !WinExist("Open patient case")
    {
        MsgBox must be on "Open patient case" to use this function
        Gui, Destroy
        Exit
    }
    
    BlockInput, MouseMove
    
        ; click the new patient button
    WinActivate Open patient case
    BlockInput MouseMove
    MouseGetPos, x, y
    Click, 78, 46
    MouseMove, %x%, %y%, 0
    BlockInput MouseMoveOff
    
    GuiControl,, Progress, 33
    
    ; see which window opens
    WinWaitActive, ahk_group ThreeShapePatient,, 5
	if ErrorLevel
    {
		BlockInput, MouseMoveOff
        Gui, Destroy
        MsgBox, Couldn't get focus patient info popup
        Exit
    }
    
    
    GuiControl,, Progress, 66
    
    SetTitleMatchMode, 3
    if WinActive("New patient info",, model)
    {
        BlockInput, MouseMove
        Ortho_NewPatientEntry()
        BlockInput, MouseMoveOff
    }
    else if WinActive("New patient model set info")
    {
        ControlFocus, TEdit5
        Sleep, 100
        Send, %scriptnumber%
    }
    else
    {
        BlockInput, MouseMoveOff
        MsgBox,, Wrong Window, Didn't land on the "New patient info" or "New model set" window
        Gui, Destroy
        Exit
    }
    
    GuiControl,, Progress, 100
    Gui, Destroy
        
    BlockInput, MouseMoveOff
	
	Log_Increment(5, 20)
	
    return
}

; Function to be used on imported case field, loops around and deletes stls
f9::
{
	if Step = Importing
	{
		
		if !WinExist("OrthoAnalyzer. Patient ID:")
		{
			MsgBox Must be in case view window in Ortho Analyzer for this function
			Exit
		}
	
		SetTitleMatchMode, 1
		WinActivate OrthoAnalyzer. Patient ID:
		WinWaitActive OrthoAnalyzer. Patient ID:
	
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, 27, 43
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
		
		Log_Increment(6, 0)
	
		if FileExist(A_MyDocuments "\Temp Models\*.stl")
			Loop, 10
				FileDelete, %A_MyDocuments%\Temp Models\*.stl
	}
	
	else if Step = Prepping
	{
		if WinExist("ahk_group ThreeShape")
		{
			Ortho_Export()
			Log_Increment(7, 2)
		}
		else
		{
			MsgBox,, Wrong Window, Must be in a case to use this function
			Exit
		}
	}
	
	else
	{
		MsgBox,, Wrong Step, Must be on the importing or prepping step for this function
	}
	
    return
}

; =========================================================================================================================
; OrthoCAD Shortcuts
; =========================================================================================================================

; Shortcut to go to cadent search page and enter the patient name into fields
f10::
{
	Cadent_StillOpen()
	
	Neo_GetInfoFromReview()
	
    ;Progress Bar instance
    Gui, Add, Progress, vprogress w300 h45
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop
    
    GuiControl,, Progress, 33
	
    Cadent_Orders()
    
    GuiControl,, Progress, 66
    
    BlockInput, MouseMove
    
    if StrLen(firstname) > 1 and StrLen(lastname) > 1
        Send %lastname%, %firstname%
    
    BlockInput, MouseMoveOff
    
    GuiControl,, Progress, 100
    Gui, Destroy
	
	Log_Increment(3, 10)

	return
}

; OrthoCad Export Hotkey
f11::
{
	global MyCadentDriver, firstname, lastname, ProgX, ProgY
	
	Cadent_StillOpen()
	
	if !InStr(MyCadentDriver.Url, "https://mycadent.com/CaseInfo.aspx") ; checks to ensure on a case page
	{
		MsgBox,, Wrong Page, Must be on a case page in MyCadent
		Gui, Destroy
		Exit
	}
	
	try MyCadentDriver.findElementByID(Path_CadentExport).click() ; click on the export button
	catch e
	{
		Gui, Destroy
		Msgbox,, Web Error, Couldn't click on the export button on MyCadent
		Exit
	}
    
    Gui, Add, Progress, vprogress w300 h45
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop
	
	BlockInput, MouseMove
	
	SetTitleMatchMode, 1
	WinWait, ahk_exe OrthoCAD.exe,, 30, Export  ; Wait for main window to open
	if ErrorLevel
		{
			Gui, Destroy
			BlockInput, MouseMoveOff
			MsgBox,, OrthoCAD Error, Main OrthoCAD window didn't open
			Exit
		}
		
	Sleep, 400 ; buffer for file to load
	Cadent_Orders()
		
	GuiControl,, Progress, 20
		
	SetTitleMatchMode, 3
	WinWait, OrthoCAD Export,, 30   ; Wait for export box to open
	if ErrorLevel
		{
			Gui, Destroy
			BlockInput, MouseMoveOff
			MsgBox,, OrthoCAD Error, Export box didn't open
			Exit
		}
		
	GuiControl,, Progress, 40
    
    ; Activate orthocad, then the export window
    WinActivate, ahk_exe OrthoCAD.exe
    WinActivate, "OrthoCAD Export"
    
    ; set the export type to open scan
    ControlFocus, ComboBox1, OrthoCAD Export
    Send, {down}
    Sleep, 100
    
    ; set the export type to models in occlusion
    ControlFocus, ComboBox2, OrthoCAD Export
    Send, {down}{down}
    Sleep, 100
    
    ; set the export folder to "export"
    ControlFocus, Edit1, OrthoCAD Export
    Send {CtrlDown}a{CtrlUp}
    Sleep, 100
    Send, %firstname%%lastname%
    Sleep, 100
    
    ; go to the export button and hit enter
    Send, {tab}
    Sleep, 100
    Send, {enter}
	
	GuiControl,, Progress, 60
	
	SetTitleMatchMode, 3
	WinWait, Export Done,, 30
	if ErrorLevel
		{
			Gui, Destroy
			BlockInput, MouseMoveOff
			MsgBox,, OrthoCAD Error, Export confirmation box didn't open
			Exit
		}
		
	WinActivate, Export Done
	WinWaitActive, Export Done,, 30
	if ErrorLevel
		{
			Gui, Destroy
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
	
	GuiControl,, Progress, 80
	
	WinKill, ahk_exe OrthoCAD.exe
	Sleep, 200
	
	SetTitleMatchMode, 2
	if WinExist("ahk_class #32770")   ; confirm closing without saving
	{
		Send {tab}
		sleep, 100
		Send {enter}
		Sleep, 200
	}
	
	WinWaitClose, ahk_exe OrthoCAD.exe,, 10
	if ErrorLevel
	{
		Gui, Destroy
		BlockInput, MouseMoveOff
		MsgBox,, OrthoCAD Error, Couldn't close OrthoCAD
		Exit
	}

	newfirst := StrReplace(firstname, " ", "")
    newlast := StrReplace(lastname, " ", "")

	IfNotExist, C:\Cadent\Export\
		FileCreateDir, C:\Cadent\Export\
	
	IfNotExist, %A_MyDocuments%\Temp Models
		FileCreateDir, %A_MyDocuments%\Temp Models
	
	IfNotExist, C:\Cadent\Export\%firstname%%lastname%
		MsgBox Couldn't find export folder, didn't move files
	
	if FileExist("C:\Cadent\Export\" newfirst newlast "\" "*u.stl")
	{
		FileMove, C:\Cadent\Export\%newfirst%%newlast%\*u.stl, %A_MyDocuments%\Temp Models\%newfirst%%newlast%upper.stl, 1
		counter += 1
	}

	if FileExist("C:\Cadent\Export\" newfirst newlast "\" "*l.stl")
	{
		FileMove, C:\Cadent\Export\%newfirst%%newlast%\*l.stl, %A_MyDocuments%\Temp Models\%newfirst%%newlast%lower.stl, 1
		counter += 1
	}
    
    if FileExist("C:\Cadent\Export\" newfirst newlast "\" "*.stl")
	{
		FileMove, C:\Cadent\Export\%newfirst%%newlast%\*.stl, %A_MyDocuments%\Temp Models\, 1
		counter += 1
	}
    
    IfExist, C:\Cadent\Export\%newfirst%%newlast%
        FileRemoveDir, C:\Cadent\Export\%newfirst%%newlast%, 1
	
	Log_Increment(8, 0)
	GuiControl,, Progress, 100
    Gui, Destroy
	BlockInput, MouseMoveOff
    return
}
    
; Finish the orthocad export    
f12::
{
    if !FileExist(A_MyDocuments "\Temp Models\*.stl")
	{
		MsgBox,, Importing Finishing Error, No STLs in the Temp Models Folder
		return
	}
	
	Temp_Models_List := [] ; Array for listing files to pull from
	Filename_Upper := ""  ; set these empty so the duplicate check works later
	Filename_Lower := ""

	Loop Files, %A_MyDocuments%\Temp Models\*.stl ; Appends each file to the list
		Temp_Models_List.Push(A_LoopFileName)
	
	if(Temp_Models_List.Length() > 2)
	{
		MsgBox,, Importing finishing Error, Too many STLs in Temp Models folder
		return
	}
	
	Tags_Upper := ["Upr.stl", "u.stl", "upper.stl", "max.stl", "Max.stl", "_u-"]
	Tags_Lower := ["Lwr.stl", "l.stl", "lower.stl", "man.stl", "Man.stl", "_l-"]
	
	for key, filename in Temp_Models_List
	{
		Pass_check := 0
		for key, tag in Tags_Upper ; all possible marking for an upper
		{
			if instr(filename, tag)
			{
				if (Filename_Upper != "")
				{
					msgbox,, Importing Finishing Error, found more than one upper STL in Temp Models folder
					return
				}
				Filename_Upper := filename
				Pass_check := 1
			}
		}
		for key, tag in Tags_Lower ; all possible marking for an upper
		{
			if instr(filename, tag)
			{
				if (Filename_Lower != "")
				{
					msgbox,, Importing Finishing Error, found more than one lower STL in Temp Models folder
					return
				}
				Filename_Lower := filename
				Pass_check := 1
			}
		}
		if (Pass_check = 0)
		{
			msgbox,, Importing Finishing Error, STLs must be labeled as upper or lower
			return
		}
	}
	
	Neo_GetInfoFromReview()
	
	clinic := StrReplace(clinic, "#", "") ; Specifically for Boston Children's Hospital cases, hash messes with auto import
	
	Filename_Base := firstname "~" lastname "~" StrReplace(clinic, " ", "_") "~" scriptnumber "~"
	
	; These two ifs for single arches
	if (Filename_Upper != "") and (Filename_Lower = "")
	{
		Gui, Add, Text, x30 y20 w300 h14 +Center, Normal Import:
		Gui, Add, Button, x12 y40 w100 h30 gUpper, Upper

		Gui, Add, Text, x30 y120 w300 h14 +Center, ABR Import:
		Gui, Add, Button, x12 y140 w100 h30 gUpperABR, Upper ABR

		Gui, Show, w439 h253, Import Type Selection (Upper Only)
		
		WinWaitClose, Import Type Selection (Upper Only)
		
		FileMove, %A_MyDocuments%\Temp Models\%Filename_Upper%, %Arch_Destination%%Filename_Base%Upr[1].stl
		if !FileExist(Arch_Destination Filename_Base "Upr[1].stl")
		{
			Msgbox,, Importing Finishing Error, Attempted to move STL, but couldn't verify it's location after move
			return
		}
	}
	else if (Filename_Upper = "") and (Filename_Lower != "")
	{
		; GUI to select ABR vs normal import
		Gui, Add, Text, x30 y20 w300 h14 +Center, Normal Import:
		Gui, Add, Button, x112 y40 w100 h30 gLower, Lower

		Gui, Add, Text, x30 y120 w300 h14 +Center, ABR Import:
		Gui, Add, Button, x112 y140 w100 h30 gLowerABR, Lower ABR

		Gui, Show, w439 h253, Import Type Selection (Lower Only)
		
		WinWaitClose, Import Type Selection (Lower Only)
		
		
		FileMove, %A_MyDocuments%\Temp Models\%Filename_Lower%, %Arch_Destination%%Filename_Base%Lwr[1].stl
		if !FileExist(Arch_Destination Filename_Base "Lwr[1].stl")
		{
			Msgbox,, Importing Finishing Error, Attempted to move STL, but couldn't verify it's location after move
			return
		}
	}
	
	; This one for both arches, have to be able to choose which one or both
	else if (Filename_Upper != "") and (Filename_Lower != "")
	{
		Gui, Add, Text, x30 y20 w300 h14 +Center, Normal Import:
		Gui, Add, Button, x12 y40 w100 h30 gUpper, Upper
		Gui, Add, Button, x112 y40 w100 h30 gLower, Lower
		Gui, Add, Button, x212 y40 w100 h30 gBoth, Both

		Gui, Add, Text, x30 y120 w300 h14 +Center, ABR Import:
		Gui, Add, Button, x12 y140 w100 h30 gUpperABR, Upper ABR
		Gui, Add, Button, x112 y140 w100 h30 gLowerABR, Lower ABR
		Gui, Add, Button, x212 y140 w100 h30 gBothABR, Both ABR

		Gui, Show, w439 h253, Arch Selection (Two Arches Detected)
		
		WinWaitClose, Arch Selection (Two Arches Detected)
		
		if (Arch_Finish = "Upper")
		{
			FileMove, %A_MyDocuments%\Temp Models\%Filename_Upper%, %Arch_Destination%%Filename_Base%Upr[1].stl
			if !FileExist(Arch_Destination Filename_Base "Upr[1].stl")
			{
				Msgbox,, Importing Finishing Error, Attempted to move STL, but couldn't verify it's location after move
				return
			}
			MsgBox,, Success, Upper arch added to importing queue
		}
		else if (Arch_Finish = "Lower")
		{
			FileMove, %A_MyDocuments%\Temp Models\%Filename_Lower%, %Arch_Destination%%Filename_Base%Lwr[1].stl
			if !FileExist(Arch_Destination Filename_Base "Lwr[1].stl")
			{
				Msgbox,, Importing Finishing Error, Attempted to move STL, but couldn't verify it's location after move
				return
			}
		}
		else if (Arch_Finish = "Both")
		{
			FileMove, %A_MyDocuments%\Temp Models\%Filename_Upper%, %Arch_Destination%%Filename_Base%Upr[2].stl
			FileMove, %A_MyDocuments%\Temp Models\%Filename_Lower%, %Arch_Destination%%Filename_Base%Lwr[2].stl

			if !FileExist(Arch_Destination Filename_Base "Upr[2].stl") or !FileExist(Arch_Destination Filename_Base "Lwr[2].stl")
			{
				Msgbox,, Importing Finishing Error, Attempted to move STL, but couldn't verify it's location after move
				return
			}
		}

	}
	
	else
	{
		msgbox no files were found for finishing
		return
	}
	
	if FileExist(A_MyDocuments "\Temp Models\*.stl")
			Loop, 10
				FileDelete, %A_MyDocuments%\Temp Models\*.stl
			
	msgbox,, Importing Complete, Successfully queued %Filename_Base% for importing
	
	return
}

Upper:
{
	Arch_Finish := "Upper"
	Arch_Destination := Import_Input_Dir
	Gui, Destroy
	return
}

Lower:
{
	Arch_Finish := "Lower"
	Arch_Destination := Import_Input_Dir 
	Gui, Destroy
	return
}

Both:
{
	Arch_Finish := "Both"
	Arch_Destination := Import_Input_Dir
	Gui, Destroy
	return
}

UpperABR:
{
	Arch_Finish := "Upper"
	Arch_Destination := ABR_Input_Dir
	Gui, Destroy
	return
}

LowerABR:
{
	Arch_Finish := "Lower"
	Arch_Destination := ABR_Input_Dir
	Gui, Destroy
	return
}

BothABR:
{
	Arch_Finish := "Both"
	Arch_Destination := ABR_Input_Dir
	Gui, Destroy
	return
}

; =========================================================================================================================
; Multi Site/Extra functions
; =========================================================================================================================

; Insert the order ID from iTero into a new note on the edit page, or take bite pic if prepping
Insert::
{
    global orderid, ProgX, ProgY, Step, Path_SaveNoteCSS, Path_UploadFileXPATH
	
	if Step = Importing ; On the importing step, grabs the iTero ID and inserts it into the notes field
	{
		Gui, Add, Progress, vprogress w300 h45
		Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
		Gui, +AlwaysOnTop
		
		Cadent_GetOrderID()
		
		GuiControl,, Progress, 10
		
		Neo_Activate()
		
		GuiControl,, Progress, 20
		
		Neo_NewNote()
		
		GuiControl,, Progress, 50
		
		Sleep, 200
		
		BlockInput, MouseMove
		
		Send iTero ID: %orderid%
		
		BlockInput, MouseMoveOff
		
		GuiControl,, Progress, 75
		
		
		;tab to the save button, css path doesn't work anymore
		Send {tab}{tab}
		
		Sleep, 200
		
		Send {enter}
		
		GuiControl,, Progress, 100
		Gui, Destroy
		
		Log_Increment(5, 2)
	}
	
	else ; If on any other step, takes bite screenshots in ortho and uploads them to the website
	{
		
		; gets first and last name to be used in the screenshot filenames
		Neo_StillOpen()
		Neo_GetInfoFromReview()
		
		if !WinExist("ahk_group ThreeShape")
		{
			MsgBox,, Wrong Window, Case must be open in OrthoAnalyzer or ApplianceDesigner
			Exit
		}

		; defines directory and destroys and recreates it to ensure it's empty
		dir= %A_MyDocuments%\Automation\Screenshots
		FileRemoveDir, %dir%, 1
		FileCreateDir, %dir%
		
		; front view screenshot
		FirstViewX := 1902 ; front view x position
		FirstViewY := 105 ; front view y
		RefVar := VarF
		Ortho_View()
		
		screenshotname := firstname lastname "Front.jpg"
		Sleep, 500
		CaptureScreen(screenshotname) ; function defined in CaptureScreen Library
		
		; left view screenshot
		FirstViewX := 1902 ; 
		FirstViewY := 177 ;
		RefVar := VarF
		Ortho_View()
		
		screenshotname := firstname lastname "Left.jpg"
		Sleep, 500
		CaptureScreen(screenshotname)	
		
		; right view screenshot
		FirstViewX := 1902 ;
		FirstViewY := 215 ; 
		RefVar := VarF
		Ortho_View()
		
		screenshotname := firstname lastname "Right.jpg"
		Sleep, 500
		CaptureScreen(screenshotname)	
		
		Neo_Activate()
		
		; tries to find the website button to upload files and click it
		try NeoDriver.findElementByXpath(Path_UploadFileXPATH).Click()
		catch e
		{
			MsgBox,, Website Error, Couldn't find the file upload button on the website
			Exit
		}
		
		WinWaitActive Open,, 5
		if ErrorLevel
		{
			MsgBox,, Website Error, File Upload window didn't open properly
			Exit
		}
		
		Sleep, 100
		
		Send % dir ; send the directory for the automation screencaps folder
		Sleep, 200
		
		Send {enter}
		Sleep, 200
		
		Send {shiftDown}{tab}{ShiftUp} ; tab into the main files window
		Sleep, 100
		
		Send {CtrlDown}a{CtrlUp} ; select all files
		Sleep, 100
		
		Send {enter} ; confirm
	}
	
	Log_Increment(20, 20)
	
    return
}

; ===========================================================================================================================
; OrthoAnalyzer/ApplianceDesigner Views, including Top(bottom), Front(back), Right(left), transparancy, and switching model
; ===========================================================================================================================

; If either OrthoAnalyzer - [\\APP03\New Scans or ApplianceDesigner - [\\APP03\New Scans are open
SettitleMatchMode, 1
#IfWinExist ahk_group ThreeShape 

; Top(bottom) View
!+t::
{
	FirstViewX := 1902 ; top view x position
    FirstViewY := 286 ; top view y
    SecondViewX := 1902 ; bottom view x
    SecondViewY := 253 ; bottom view y
	RefVar := VarT
	
    Ortho_View()
	
	VarT := A_TickCount
	Log_Increment(2, 0)
	return
}

; Right(left) View
!+r::
{
	FirstViewX := 1902 ; Right view x position
    FirstViewY := 215 ; right view y
    SecondViewX := 1902 ; left view x
    SecondViewY := 177 ; left view y
	RefVar := VarR
	
    Ortho_View()
	
	VarR := A_TickCount
	Log_Increment(2, 0)
	return
}

; Front(back) View
!+f::
{
	FirstViewX := 1902 ; front view x position
    FirstViewY := 106 ; front view y
    SecondViewX := 1902 ; back view x
    SecondViewY := 140 ; back view y
	RefVar := VarF
	
    Ortho_View()
	
	VarF := A_TickCount
	Log_Increment(2, 0)
	return
}

; Transparency
!+c::
{
	FirstViewX := 1902 ; transparency x position
    FirstViewY := 800 ; transparency y
    SecondViewX := 1902 ; transparency x
    SecondViewY := 800 ; transparency y
	
    Ortho_View()
	Log_Increment(2, 0)
	return
}

; Switch Visible Model
!+v::
{
	UpperDotX := 1708
	UpperDotY := 79
	LowerDotX := 1708
	LowerDotY := 106
	
	Ortho_VisibleModel()
	Log_Increment(2, 0)
	return
}

; Export finished case
!+.::
{
	Ortho_Export()
	Log_Increment(7, 2)
	return
}

; ===========================================================================================================================
; OrthoAnalyzer/ApplianceDesigner functions for pushing forward the dialogue boxes, all using !+g
; ===========================================================================================================================

SettitleMatchMode, 1
#IfWinExist ahk_group ThreeShape 

!+g::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	
	if PrepStep in %3ShapeSteps%
		WinActivate ahk_group ThreeShape
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, 190, 30
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
	return
}

; ===========================================================================================================================
; OrthoAnalyzer/ApplianceDesigner Sculpting Tool Selections
; ===========================================================================================================================

#IfWinExist ahk_group ThreeShape 

; Wax Knife preset double tap tools
!+1::
{
	FirstKnife := 1
	SecondKnife := 5
	
	RefVar := Var1
	
	Ortho_Wax()
	
	Var1 := A_TickCount
	Log_Increment(2, 0)

	return
}

!+2::
{
	FirstKnife := 2
	SecondKnife := 6
	
	RefVar := Var2
	
	Ortho_Wax()
	
	Var2 := A_TickCount
	Log_Increment(2, 0)

	return
}

!+3::
{
	FirstKnife := 3
	SecondKnife := 7
	
	RefVar := Var3
	
	Ortho_Wax()
	
	Var3 := A_TickCount
	Log_Increment(2, 0)

	return
}

!+5::
{
	FirstKnife := 4
	SecondKnife := 4
	
	RefVar := 0
	
	Ortho_Wax()
	Log_Increment(2, 0)

	return
}

; Artifact Removal
!+a::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
		return
	
	WinActivate ahk_group ThreeShape
	BlockInput MouseMove
	MouseGetPos, x, y
	Click, 120, 173
	MouseMove, %x%, %y%, 0
	BlockInput MouseMoveOff
	Log_Increment(2, 0)
	return
}

; Plane Cut
!+p::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
		return
	
	WinActivate ahk_group ThreeShape
	BlockInput MouseMove
	MouseGetPos, x, y
	Click, 165, 175
	MouseMove, %x%, %y%, 0
	BlockInput MouseMoveOff
	Log_Increment(2, 0)
	return
}

; Spline Cut
!+s::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
		return
	
	WinActivate ahk_group ThreeShape
	BlockInput MouseMove
	MouseGetPos, x, y
	Click, 205, 175
	Click, 185, 134
	MouseMove, %x%, %y%, 0
	BlockInput MouseMoveOff
	Log_Increment(2, 0)
	return
}

; ===========================================================================================================================
; 3Shape Function Library
; ===========================================================================================================================

; ---------------------------------------------------------------------------------------------------------------------------
; Import Functions

Ortho_AdvSearch() ; function to enter patient name into advanced search field and search using globals
{
    global
       
    if StrLen(firstname) < 2 or StrLen(lastname) < 2
    {
        BlockInput MouseMoveOff
        MsgBox No patient info saved from RXWizard
        Exit
    }  
   
    if !WinExist("Open patient case")
    {
        BlockInput MouseMoveOff
        MsgBox Must be in the "Open Patient Case" Dialogue for this hotkey
        Exit
    }
        
    ; GUI progress bar 
    Gui, Add, Progress, vprogress w300 h45
    Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop

    WinActivate, ahk_exe OrthoAnalyzer.exe
    WinWaitActive, Open patient case,, 10
        
    if ErrorLevel
    {
        Gui, Destroy
        MsgBox, Couldn't get focus on Ortho Analyzer
        return
    }
    Sleep, 200
    
    GuiControl,, Progress, 20
    
    ControlFocus, Advanced search, Open patient case
    Sleep, 400
    Send, {Enter}
    Sleep, 400
    
    GuiControl,, Progress, 30
    
    ControlFocus, TEdit10, Open patient case    
    Sleep, 100
    Send, %firstname%
    Sleep, 100
    
    GuiControl,, Progress, 40
     
    ControlFocus, TEdit9, Open patient case    
    Sleep, 100
    Send, %lastname%
    Sleep, 100
    
    GuiControl,, Progress, 50

    ControlFocus, Edit1, Open patient case    
    Sleep, 100
    Send, %clinic%
    Sleep, 100
    
    GuiControl,, Progress, 60

    Sleep, 500   ; added to prevent search glitch
        
    ControlFocus, TButton2, Open patient case    
    Send {enter}
    Sleep, 100
    Send {down}
    
    GuiControl,, Progress, 100
    Gui, Destroy
    
    BlockInput MouseMoveOff
    return
}

Ortho_NewPatientEntry() ; function to paste patient info into new patient info fields using globals
{
    ; paste info
    global firstname, lastname, clinic
    
    if StrLen(firstname) < 2 or StrLen(lastname) < 2
    {
        BlockInput MouseMoveOff
        MsgBox No patient info saved from RXWizard
        Gui, Destroy
        Exit
    }
    
    ControlFocus, TEdit4, New patient info
    Sleep, 200
    Send, %firstname% %lastname%
    Sleep, 100
    
    ControlFocus, TEdit10, New patient info
    Sleep, 100
    Send, %firstname%
    Sleep, 100
    
    ControlFocus, TEdit9, New patient info
    Sleep, 100
    Send, %lastname%
    Sleep, 100

    ControlFocus, Edit1, New patient info
    Sleep, 200
    Send, %clinic%

    BlockInput MouseMoveOff
    return
}

; ------------------------------------------------------------------------------------------------------------------------------
; Prepping Functions



Ortho_OpenCase()
{
	global
	
	if !WinExist("Open patient case")
    {
		BlockInput, MouseMoveOff
        MsgBox must be on "Open patient case" to use this function
        Gui, Destroy
        Exit
    }
	
	if StrLen(firstname) < 2 or StrLen(lastname) < 2
    {
        BlockInput MouseMoveOff
        MsgBox No patient info saved from RXWizard
        Exit
    }  
	
	if WinExist("Exported items")
    {
		WinActivate, Exported items
		WinWaitActive, Exported items,, 10
		ControlFocus, TButton1, Exported items
		Sleep, 300
		Send {Enter}
    
    }
	
	WinActivate, ahk_group ThreeShapeExe
    WinWaitActive, Open patient case,, 10
        
    if ErrorLevel
    {
		BlockInput, MouseMoveOff
        Gui, Destroy
        MsgBox, Couldn't get focus on Ortho Analyzer
        Exit
    }
	
	ControlFocus, Advanced search, Open patient case
    Sleep, 600
    Send, {Enter}
    Sleep, 400
	
	ControlFocus, TEdit16, Open patient case    
    Sleep, 400
    Send, %scriptnumber%
    Sleep, 400
	
	Send, {Enter}
	
	Sleep, 400
	
	Send {down}{right}{down}
	
	return
}

Ortho_View() ; clicks the view button declared before the function is called, swaps to a secondary view if pressed twice
{
    global toggle, RefVar, FirstViewX, FirstViewY, SecondViewX, SecondViewY
       
    startvar := A_TickCount
	if startvar - RefVar > 1000
	{
		WinActivate ahk_group ThreeShape
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, %FirstViewX%, %FirstViewY%
		MouseMove, %x%, %y%, 0
		toggle := 1
		BlockInput MouseMoveOff
	}
	else
	{
		if toggle = 1
		{
			WinActivate ahk_group ThreeShape
			BlockInput MouseMove
			MouseGetPos, x, y
			Click, %SecondViewX%, %SecondViewY%
			MouseMove, %x%, %y%, 0
			toggle := 2
			BlockInput MouseMoveOff
		}
		else if toggle = 2
		{
			WinActivate ahk_group ThreeShape
			BlockInput MouseMove
			MouseGetPos, x, y
			Click, %FirstViewX%, %FirstViewY%
			MouseMove, %x%, %y%, 0
			toggle := 1
			BlockInput MouseMoveOff
		}
	}
	return
}

Ortho_VisibleModel() ; Swaps Visisble Model
{
    global UpperDotX, UpperDotY, LowerDotX, LowerDotY
    
    ; Activates the main window and gets the pixel colors of the model dots
	WinActivate ahk_group ThreeShape
	PixelGetColor, uppervar, %UpperDotX%, %UpperDotY%
	PixelGetColor, lowervar, %LowerDotX%, %LowerDotY%
    
    ; if both black, turn off the lower
	if (uppervar = 0x000000) and (lowervar = 0x000000)
	{
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, %LowerDotX%, %LowerDotY%
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
	}
    ; if lower off, switch models
	if (uppervar = 0x000000) and (lowervar = 0xF0F0F0)
	{
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, %UpperDotX%, %UpperDotY%
		Click, %LowerDotX%, %LowerDotY%
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
	}
    ; if upper off, turn on upper
	if (uppervar = 0xF0F0F0) and (lowervar = 0x000000)
	{
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, %UpperDotX%, %UpperDotY%
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
	}
	return
}

Ortho_Wax() ; Tool for swapping out wax knifes
{
    global RefVar, FirstKnife, SecondKnife
    
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	ControlGetText, PrepTool, TdfGroupInfo1, ahk_group ThreeShape
	
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
		return
		
	if PrepTool != "Wax knife settings"
		WinActivate ahk_group ThreeShape
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, 35, 175
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
	startvar := A_TickCount
	if startvar - RefVar > 900
	{
		Send %FirstKnife%
	}
	else
	{
		Send %SecondKnife%
	}
	return
}

Ortho_Export() ; While in model edit, clicks the green check mark and exports model
{
    global

    Gui, Add, Progress, vprogress w300 h45
    Gui, +AlwaysOnTop
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running

    
	; Check Mark
	BlockInput MouseMove
	WinActivate ahk_group ThreeShape
    
    GuiControl,, Progress, 10
    
    ; Click Green Check Mark
	Click, 1088, 95
	BlockInput MouseMoveOff
    
    GuiControl,, Progress, 20
    
	WinWaitActive, Open patient case,, 30
	if ErrorLevel
	{
		BlockInput MouseMoveOff
        Gui, Destroy
		MsgBox,, Timeout Error, "Open Patient Case" window took too long to open
		Exit
	}
    
    GuiControl,, Progress, 40
    
	; Export Clicks
	BlockInput MouseMove
	Sleep, 200
	Click, 389, 46
	Sleep, 100
	Send {Down 6}
	Sleep, 100
	Send {Enter}
    
    GuiControl,, Progress, 50
    
    ; Wait for export confirmation box
	WinWaitActive, Exported items,, 15
	if ErrorLevel
	{
		BlockInput MouseMoveOff
        Gui, Destroy
		MsgBox,, Timeout Error, Export Confirmation took too long
		Exit
	}
    
    GuiControl,, Progress, 60
    
	
	BlockInput MouseMoveOff
    
    GuiControl,, Progress, 100
    Gui, Destroy
    
	return
}

; ===========================================================================================================================
; Library for all Standalone functions pertaining to Netfabb
; ===========================================================================================================================

Netfabb_Level() ; starts leveling, finishes leveling after the click, and removes the max/man tags in engraving
{
	global
	
	If !WinActive(ahk_exe netfabb.exe "Autodesk Netfabb") ; if netfabb isn't the active window, display error and stop
	{
		Msgbox,, Wrong Program Active, F1 should only be used while Netfabb is active
		Exit
	}
	
	; Level Model
	Send {Control Down}w{Control Up}
	Sleep, 100
	
	if WinExist(ahk_exe netfabb.exe "Warning")
	{
		Exit
	}
	
	KeyWait, LButton, D ; waits for the user to click the bottom of the model
	BlockInput, MouseMove
	
	Gui, Add, Progress, vprogress w300 h45 ; progress bar gui
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop	
	
	Sleep, 300
	ControlFocus, Apply, Autodesk Netfabb ; hits the apply button on the leveling step
	Sleep, 100
	Send {Enter}
	
	GuiControl,, Progress, 20
	
	Sleep, 200
	
	Send {Control Down}e{Control Up} ; shortcut for the engrave function
	Sleep, 200
	
	GuiControl,, Progress, 40
	
	; Remove Man/Max, could exist in any of these 4 edit fields, depending on the install version
	ControlGetText, Engraving4, Edit4, Autodesk Netfabb
	ControlGetText, Engraving17, Edit17, Autodesk Netfabb
	ControlGetText, Engraving18, Edit18, Autodesk Netfabb
	ControlGetText, Engraving1, Edit1, Autodesk Netfabb
	ControlGetText, Engraving15, Edit15, Autodesk Netfabb
	
	GuiControl,, Progress, 60
	
	If InStr(Engraving17, "Max") or InStr(Engraving17, "Man")
	{
		ControlFocus, Edit17, Autodesk Netfabb
		Netfabb_DeleteTag() ; see this function at the end, just consolidates the number of lines here.
	}
	If InStr(Engraving4, "Max") or InStr(Engraving4, "Man")
	{
		ControlFocus, Edit4, Autodesk Netfabb
		Netfabb_DeleteTag()
	}
		If InStr(Engraving18, "Max") or InStr(Engraving18, "Man")
	{
		ControlFocus, Edit18, Autodesk Netfabb
		Netfabb_DeleteTag()
	}
	If InStr(Engraving1, "Max") or InStr(Engraving1, "Man")
	{
		ControlFocus, Edit1, Autodesk Netfabb
		Netfabb_DeleteTag()
	}
	If InStr(Engraving15, "Max") or InStr(Engraving15, "Man")
	{
		ControlFocus, Edit15, Autodesk Netfabb
		Netfabb_DeleteTag() ; see this function at the end, just consolidates the number of lines here.
	}
	
	else ; if there was no man or max, just wait until the engraving click to update the confirmation variable
	{
		Gui, Destroy
		BlockInput MouseMoveOff
		KeyWait, LButton, D
		placedengraving =: 1
	}
	return
}

Netfabb_Finish() ; function to finish engraving, and attempt to export the model. won't pass the shell error message box
{
	global
	
	If !WinActive(ahk_exe netfabb.exe "Autodesk Netfabb") ; confirms that netfabb is the active window
	{
		MsgBox,, Wrong Window Active, F2 should only be used in Netfabb, or to rename in Windows Explorer
		Exit
	}
	
	; tries to find the check box for "keep aspect ratio", which should only appear in the engraving module
	try ControlGetText, aspectratio, Keep Aspect Ratio, ahk_exe netfabb.exe 
	catch e
	{
		Msgbox,, Wrong Step, F2 should only be used immediately after F1, after placing the engraving (couldn't find aspect button)
		Exit
	}

			
	If placedengraving =: 1 ; variable from F1 function that only appears if there was a click to place the engraving
	{
		; Finish Engrave
		BlockInput, MouseMove
		
		; progress bar gui
		Gui, Add, Progress, vprogress w300 h45
		Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
		Gui, +AlwaysOnTop		
		
		; added wait to attempt to fix error with missing the control focus send
		WinActivate, Autodesk Netfabb
		WinWaitActive, Autodesk Netfabb
		
		ControlFocus, Apply, Autodesk Netfabb ; hits the apply button for the engraving
		
		Sleep, 100 ; Another wait to attempt to fix the wait error		
		Send {Enter}
		
		GuiControl,, Progress, 20
		
		WinWaitActive, Confirmation,, 2 ; confirms the "delete original model" popup
		Send {Enter}
		
		GuiControl,, Progress, 40
		
		placedengraving =: 0 ; resets this counter for next time
		
		Sleep, 2000 ; waits for the engraving to finish
		
		GuiControl,, Progress, 60
		
		; Export
		Send !f
		Sleep, 300
		Send r
		Sleep, 100
		Send {Down 2}{Enter} ; gets down to "as STL"
		WinWaitActive, Export,, 1
		
		GuiControl,, Progress, 80
		
		if ErrorLevel ; if the export window couldn't be found
		{
			BlockInput MouseMoveOff
			Gui, Destroy
			MsgBox Timeout
			return
		}
		else
			Send {Enter 2} ; once to confirm, a second time to hit the "optimize" button if it appears
			Gui, Destroy
			BlockInput MouseMoveOff
		return
	}
	else
	{
		MsgBox Click to place engraving first
		return
	}
}

Netfabb_Loop() ; deletes the highlighted model and returns to import, or repairs model if on the export box with an error
{
	global
	; if the main box is up, delete the part and return to import
	If WinActive(ahk_exe netfabb.exe "Autodesk Netfabb")
	{
		BlockInput, MouseMove
		
		; progress bar gui
		Gui, Add, Progress, vprogress w300 h45
		Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
		Gui, +AlwaysOnTop	

		WinActivate, ahk_exe netfabb.exe
		
		sleep, 200
		
		; Circle Around
		Send {Delete}{Enter}
		
		GuiControl,, Progress, 33
		
		Sleep, 100
		Send !f
		
		GuiControl,, Progress, 66
		
		Sleep, 100
		Send p
		
		GuiControl,, Progress, 100
		Gui, Destroy
		BlockInput, MouseMoveOff
		
		return
	}
	
	; if the dialogue box is up, run the repair script
	else if WinActive(ahk_exe netfabb.exe "Export")
	{
		BlockInput MouseMove
		; progress bar gui
		Gui, Add, Progress, vprogress w300 h45
		Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
		Gui, +AlwaysOnTop	
		
		; Repair
		ControlFocus, Cancel, Export ; hits the cancel button on the export
		Send {Enter}
		Sleep, 100
		
		GuiControl,, Progress, 20
		
		Send !pe ; gets to the repair module
		Sleep, 200
		
		GuiControl,, Progress, 40
		
		Send {Down 2}`t{Enter}
		WinWait, Autodesk Netfabb Standard, One job in queue ; waits for the repair job to enter que
		Sleep, 1000
		
		GuiControl,, Progress, 60
		
		WinWait, Autodesk Netfabb Standard, No jobs ; and waits for it to finish
		Send `t`t`t{Down}
		
		GuiControl,, Progress, 80
		
		; Export
		Send !f
		Sleep, 100
		
		GuiControl,, Progress, 90
		
		Send r
		Sleep, 100
		
		GuiControl,, Progress, 100
		
		Send {Down 2}{Enter}
		WinWaitActive, Export,, 1
		if ErrorLevel
		{
			BlockInput MouseMoveOff
			Gui, Destroy
			MsgBox Timeout
			return
		}
		else
			Send {Enter 2}
			BlockInput MouseMoveOff
			Gui, Destroy
		return
	}

	else
	{
		MsgBox,, Wrong Window, F3 should only be used in NetFabb
		Exit
	}
}

Netfabb_DeleteTag() ; used at the end of F1, deletes the man or max tag off the end
{
	global
	Sleep, 100
	Send {End}
	Sleep, 100
	Send `b`b`b`b
	Gui, Destroy
	BlockInput MouseMoveOff
	KeyWait, LButton, D
	placedengraving =: 1
}

; ==========================================================================================================================
; Web Library, all website subfunctions live here
; ==========================================================================================================================

; Don't Touch, attaches web driver to last opened tab
ChromeGet(IP_Port := "127.0.0.1:9222") 
	{
		Driver := ComObjCreate("Selenium.ChromeDriver")
		Driver.SetCapability("debuggerAddress", IP_Port)
		Driver.Start()
		return Driver
	}

; -----------------------------------------------------------------------------------------------------------------------
; RXWizard functions

Neo_StartWebDriver()
{
	global
	
	if WinExist("ahk_exe chrome.exe")
		MsgBox, 4, Close Chrome?, Close all current instances of Chrome? (suggested)
			IfMsgBox, Yes
				; Close all instances of chrome
				While WinExist("ahk_exe chrome.exe")
				{
					Loop, 10
					{
						WinClose, ahk_exe chrome.exe
					}
				}	
			IfMsgBox, No
				nothing := 0
		
	; Check to make sure that the modified Chrome shortcut is in the right place
	chromelocation = %A_MyDocuments%\Automation\ChromeForAHK.lnk
	if !FileExist(chromelocation)
	{
		MsgBox,, Chrome Shortcut Not Found, Can't find the proper shortcut at %A_MyDocuments%\Automation\ChromeForAHK.lnk `nRemember to modify the shortcut to run in debug mode by adding "--remote-debugging-port=9222"
		Exit
	}

	; Open Chrome instance and assign it to neolab
	Run, ChromeForAHK.lnk, %A_MyDocuments%\Automation\
	Sleep, 500
	NeoDriver := ChromeGet()
	NeoDriver.Get("https://portal.rxwizard.com/cases")
	
	NeoCheck := 1
	
	Exit
}

; function to check that the NeoDriver is still open
Neo_StillOpen()
{
	global NeoDriver, NeoCheck
	
	if NeoCheck = 0
		MsgBox, 4, Open NeoDriver?, No instance of NeoDriver found, initiate?
			IfMsgBox, Yes
				Neo_StartWebDriver()
			IfMsgBox, No
				Exit
	
	try temp := InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	catch e
	{
		MsgBox, 4, Webdriver Error, The tab for driving Portal.RXWizard was closed, initiate webdriver?
			IfMsgBox, Yes
				Neo_StartWebDriver()
			IfMsgBox, No
			{
				Gui, Destroy
				Exit
			}
		Gui, Destroy		
		return
	}
}
	
; function to bring NEOLab website to the front
Neo_Activate()
{
	if WinExist("New England Orthodontic Laboratory - Google Chrome")
		WinActivate
	else
	{
		MsgBox,, Website Error, RxWizard Portal should be the top tab in its own instance
		exit
	}
	return
}
	
; function for returning to the cases page and entering the search field
Neo_NavigateToCases()
{
    global NeoDriver, Path_ScanScriptCSS

	Neo_StillOpen()

    if !InStr(NeoDriver.Url, "https://portal.rxwizard.com") ; Must be on a rxwizard page
    {      
		Gui, Destroy
        MsgBox,, Wrong Page, Must be on an RxWizard Page
        Exit
    }     
	
	Neo_Activate()
	
	Send {tab}
	Sleep, 100
    NeoDriver.findElementByCss(Path_ScanScriptCSS).click()

}

; function for getting to the review page from the edit page
Neo_NavigateReviewFromEdit()
{
    global NeoDriver, Path_ReviewButtonCSS

	Neo_StillOpen()
	
	Neo_Activate()
	
	; If on the edit page, click on the "review" button on the bottom, then either wait for the review page or
	; click the "confirm" button on the "over model count" popup.
    if InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	{
        try NeoDriver.findElementByCss(Path_ReviewButtonCSS).sendKeys(Keys.RETURN)
		catch e 
		{
			MsgBox,, Couldn't Find Element, Couldn't find the "Review" button
			Gui, Destroy
			Exit
		}
		
	}
	else if InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")
		do = 0
	else
		MsgBox,, Wrong Page, Must be on the edit page for this function
	return

}

; function for hitting start/stop for %step% while already on the review page
Neo_Start2()
{
	global NeoDriver, currentstep, currentstepxpath
		
	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/") ; checks to ensure on the review page
	{
		MsgBox,, Wrong Page, Must be on the review page for this function
		Gui, Destroy
		Exit
	}
	
	; Wait until the button for pushing step is clickable
	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.CSS_SELECTOR, currentstepxpath))
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the start button before it was clicked the first time
		Exit
	}
		
	; Try to find the text label on the button for pushing steps
	try webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't get the text from the start stop button before starting
		Exit
	}
	
	; check the retrieved button text against the currenstep variable
	if webstep = %currentstep%
	{
		try NeoDriver.findElementByCss(currentstepxpath).Click
		catch e
		{
			Gui, Destroy
			Msgbox,, Web Error, Couldn't click on the button at the start step
			Exit
		}
		
		Neo_Activate()
		
		WinWaitActive, New England Orthodontic Laboratory - Google Chrome,, 10
		
		Sleep, 200
		
		Send, {Enter}
	}
	else
	{
		Gui, Destroy
		MsgBox,, Web Error, Button wasn't on the right step
		Exit
	}
	return
}

; new function for stopping step, will replace the original
Neo_Stop2()
{
	global NeoDriver, currentstep, currentstepxpath
	
	; Wait until the button for pushing step is clickable
	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.CSS_SELECTOR, currentstepxpath))
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the stop button before it was clicked 
		Exit
	}
	
	; Try to find the text label on the button for pushing steps
	try webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't get the text from the stop button before stopping
		Exit
	}
	
	; loop to check if the button is the right step, clicks it if it is, returns for 5 200ms loops if not
	failcount := 0
	Loop
	{
		if webstep = %currentstep%
		{
			NeoDriver.findElementByCss(currentstepxpath).Click
						
			Neo_Activate()
			
			WinWaitActive, New England Orthodontic Laboratory - Google Chrome,, 10
			
			Sleep, 200
			
			Send, {Enter}
			break
		}
		else if failcount < 20
		{
			Sleep, 400
			webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
			failcount++
			continue
		}
		else
		{
			Gui, Destroy
			MsgBox,, Website Error, Button didn't update with current step
			Exit
		}
	}
	
	try webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't get the text from the stop button after stopping
		Exit
	}
	
	failcount := 0
	Loop
	{
		if webstep != %currentstep%
		{
			break
		}
		else if failcount < 20
		{
			Sleep, 400
			webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
			failcount++
			continue
		}
		else
		{
			Gui, Destroy
			MsgBox,, Website Error, Button didn't update with current step
			Exit
		}
	}
	return
}

; function to retrieve patient and clinic names from the edit page
Neo_GetInfoFromReview()
{
    global NeoDriver, firstname, lastname, clinic, scriptnumber, Path_ScriptNumberCSS, Path_ClinicNameCSS, Path_PatientNameCSS
	Neo_StillOpen()
	
	Neo_Activate()
	
    if InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")
    {
        ; get the script from the top of the review page
        try scriptnumber := NeoDriver.findElementByCss(Path_ScriptNumberCSS).Attribute("innerText")
		catch e
			MsgBox Couldn't find the script number
		
        scriptnumber := StrReplace(scriptnumber, "Case ")
        
		; attempt to get the first, last, and clinic.
        try
		{
			fullname := NeoDriver.findElementByCss(Path_PatientNameCSS).Attribute("innerText")
			clinic := NeoDriver.findElementByCss(Path_ClinicNameCSS).Attribute("innerText")
		}
		catch e
		{
			MsgBox, couldn't find patient or clinic name
			Gui, Destroy
			Exit
		}
		
		; take invalid characters out of the names
		invalidchars := ["...", "..", ".", ",", "'"]
		for item in invalidchars
			fullname := StrReplace(fullname, invalidchars[item], "")
		
		firstname := StrSplit(fullname, " ")[1]
		lastname := StrReplace(fullname, firstname " ", "")

    }
    else
	{
        MsgBox Not on the review page
		Exit
	}
    return
}

; Puts new note onto the edit page
Neo_NewNote()
{
	global NeoDriver, orderid, Path_NewNoteCSS
	; Checks the URL to make sure it's on an edit page

	Neo_StillOpen()

	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")
	{
		MsgBox Must be on review page to use this function
		Gui, Destroy
		Exit
	}
	
	; New Note Button
	try NeoDriver.findElementByCss(Path_NewNoteCSS).Click()
	catch e 
	{
		MsgBox,, Couldn't Find Element, Couldn't find the "new note" button
		Gui, Destroy
		Exit
	}
	
	BlockInput, MouseMove
	
	Sleep, 200
	
	; Dropdown for input type
	try NeoDriver.findElementByID("note_type").Click()
	catch e 
	{
		BlockInput, MouseMoveOff
		MsgBox,, Couldn't Find Element, Couldn't find the note type drop down
		Gui, Destroy
		Exit
	}
	
	Sleep, 200
	
	; Select CSR
	Send {enter}
	
	Sleep, 300
	
	; enter note field
	
	try NeoDriver.findElementByID("note_text").Click()
	catch e 
	{
		BlockInput, MouseMoveOff
		MsgBox,, Couldn't Find Element, Couldn't find the note text box
		Gui, Destroy
		Exit
	}
	
	return
}
	
; --------------------------------------------------------------------------------------------------------------------------
; MyCadent Functions

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
	
	try temp := InStr(MyCadentDriver.Url, "https://portal.rxwizard.com/cases/edit/")
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
}

; function to go to orders page
Cadent_Orders()
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
	
	return
}
	
Cadent_GetOrderID()
{
	global MyCadentDriver, orderid, Path_CadentOrderNumberID
	
	Cadent_StillOpen()
	
	if !InStr(MyCadentDriver.Url, "https://mycadent.com/CaseInfo.aspx")
	{
		MsgBox Must be on MyCadent order page to use this function
		Gui, Destroy
		Exit
	}
	
	orderid := MyCadentDriver.findElementByID(Path_CadentOrderNumberID).Attribute("value")
	
	return
	
}

; ===========================================================================================================================
; Library for all functions pertaining to Logging
; ===========================================================================================================================

; start variables for logging numbers for the server txt, will reset to zero every time they're logged
importedserver = 0
preppedserver = 0
engravedserver = 0
hotkeysserver = 0
clicksserver = 0
strokesserver = 0

; start variables for logging numbers for the local txt, will continue on counting after importing from local log
importedlocal = 0
preppedlocal = 0
engravedlocal = 0
hotkeyslocal = 0
clickslocal = 0
strokeslocal = 0

Log_ImportLocalLog() ; Imports the current log count from the local log, creates a log if none exists
{
   global
   LocalLogDir := A_MyDocuments "\Automation\Logs\" ;folder for logs
   LocalLogFile := A_MyDocuments "\Automation\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" ; log file, named with todays date

   ; sets the text in case it has to create a new txt file
   LocalLogText := Log_GenerateText(importedlocal, preppedlocal, engravedlocal, hotkeyslocal, clickslocal, strokeslocal) 
   
   if !FileExist(LocalLogDir)
   {
      FileCreateDir, %LocalLogDir%
   }
   if !FileExist(LocalLogFile)
   {
      FileAppend, %LocalLogText%, %LocalLogfile% ; creates a new log file with the current variables
   }
   
   FileRead, ExtantLog, %LocalLogfile%
   
   LogArray := StrSplit(ExtantLog, "`n") ; creates an array of the individual lines of the txt
   
   ; for each variable, remove the text and convert the number string to a integer
   importedlocal := GetNumber(LogArray[1], "Cases imported = ")
   preppedlocal := GetNumber(LogArray[2], "Cases prepped = ")
   engravedlocal := GetNumber(LogArray[3], "Cases engraved = ")
   hotkeyslocal := GetNumber(LogArray[4], "Hotkeys used = ")
   clickslocal := GetNumber(LogArray[5], "Mouse actions saved = ")
   strokeslocal := GetNumber(LogArray[6], "Keystrokes saved = ")
   
   return
}

; function to be called in main program, increments the log by (#clicks, #strokes), and 1 hotkeysused
Log_Increment(clicks = 0, strokes = 0)
{
   global
   
   clickslocal := clicks + clickslocal
   clicksserver := clicks + clicksserver
   
   strokeslocal := strokes + strokeslocal
   strokesserver := strokes + strokesserver   
   
   hotkeyslocal+=1
   hotkeysserver+=1
   
   return
}

; updates the log txt with the current log variables plus existing txt variables, then sets local variables to 0
Log_UpdateServer()
{
   global
   
   ServerLogDir := "\\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Logs\" ;folder for logs
   ServerLogFile := "\\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" 

   ; sets the text in case it has to create a new txt file
   ServerLogText := Log_GenerateText(importedserver, preppedserver, engravedserver, hotkeysserver, clicksserver, strokesserver) 
   
   if !FileExist(ServerLogDir)
   {
      FileCreateDir, %ServerLogDir%
   }
   if !FileExist(ServerLogFile)
   {
      FileAppend, %ServerLogText%, %ServerLogfile% ; creates a new log file with the current variables
   }
   
   FileRead, ExtantLog, %ServerLogfile%
   
   LogArray := StrSplit(ExtantLog, "`n") ; creates an array of the individual lines of the txt
   
   ; for each variable, remove the text and convert the number string to a integer
   importadd := GetNumber(LogArray[1], "Cases imported = ")
   prepadd := GetNumber(LogArray[2], "Cases prepped = ")
   engraveadd := GetNumber(LogArray[3], "Cases engraved = ")
   hotkeysadd := GetNumber(LogArray[4], "Hotkeys used = ")
   clicksadd := GetNumber(LogArray[5], "Mouse actions saved = ")
   keystrokesadd := GetNumber(LogArray[6], "Keystrokes saved = ")
   
   ; adds 1 to the corresponding current step
   if Step = Importing
   {
      importedserver = 1
   }
   else if Step = Prepping
   {
      preppedserver = 1
   }
   else if Step = Engraving
   {
      engravedserver = 1
   }

   ; adds the local log variables with the corresponding variables from the txt
   importedserver := importedserver + importadd
   preppedserver := preppedserver + prepadd
   engravedserver := engravedserver + engraveadd
   hotkeysserver := hotkeysserver + hotkeysadd
   clicksserver := clicksserver + clicksadd
   strokesserver := strokesserver + keystrokesadd
   
   ; updates the log text, then writes to the file
   ServerLogText := Log_GenerateText(importedserver, preppedserver, engravedserver, hotkeysserver, clicksserver, strokesserver)
   FileDelete, %ServerLogFile%
   FileAppend, %ServerLogText%, %ServerLogfile%
   
   ; reset server variables to 0
   importedserver = 0
   preppedserver = 0
   engravedserver = 0
   hotkeysserver = 0
   clicksserver = 0
   strokesserver = 0

   return
}

Log_UpdateLocal()
{
   global
   
   ; just in case it needs to update to todays date
   LocalLogFile := A_MyDocuments "\Automation\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" ; log file, named with todays date
   
   if Step = Importing
   {
      importedlocal += 1
   }
   else if Step = Prepping
   {
      preppedlocal += 1
   }
   else if Step = Engraving
   {
      engravedlocal += 1
   }
   
   LocalLogText := Log_GenerateText(importedlocal, preppedlocal, engravedlocal, hotkeyslocal, clickslocal, strokeslocal)
   FileDelete, %LocalLogFile%
   FileAppend, %LocalLogText%, %LocalLogfile%
   return
}

; generated the text to be put into the txt file, based on the variables passed
Log_GenerateText(imported = 0, prepped = 0, engraved = 0, hotkeys = 0, clicks = 0, strokes = 0)
{
   LogText = Cases imported = %imported%`nCases prepped = %prepped%`nCases engraved = %engraved%`nHotkeys used = %hotkeys%`nMouse actions saved = %clicks%`nKeystrokes saved = %strokes%
   return LogText
}

; function to get the number from the individual string from the txt log
GetNumber(x = "", y = "")
{
   z := StrReplace(x, y, "") ; removes the leading text
   z := StrReplace(z, "`n", "") ; removes the new line markers
   z += 0 ; makes the number an integer
   return z
}

; =========================================================================================================================
; bindings for selecting model to import
; =========================================================================================================================

#IfWinActive Arch Selection

Up::
{
	gosub Upper
	return
}

Down::
{
	gosub Lower
	return
}
	
Right::
{
	gosub Both
	return
}

; =========================================================================================================================
; GIF easter egg
; =========================================================================================================================
#IfWinActive

^u::
{
	if GetKeyState("f") = 1
	{
	
		AnimPicFile := "\\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Files\JP.gif" 
		Gui, +ToolWindow
		AGif := AddAnimatedGIF(AnimPicFile)
		Gui, Show
		return

	}
}

AddAnimatedGIF(imagefullpath , x="", y="", w="", h="", guiname = "1")
{
	global AG1,AG2,AG3,AG4,AG5,AG6,AG7,AG8,AG9,AG10
	static AGcount:=0, pic
	AGcount++
	html := "<html><body style='background-color: transparent' style='overflow:hidden' leftmargin='0' topmargin='0'><img src='" imagefullpath "' width=" w " height=" h " border=0 padding=0></body></html>"
	Gui, AnimGifxx:Add, Picture, vpic, %imagefullpath%
	GuiControlGet, pic, AnimGifxx:Pos
	Gui, AnimGifxx:Destroy
	Gui, %guiname%:Add, ActiveX, % (x = "" ? " " : " x" x ) . (y = "" ? " " : " y" y ) . (w = "" ? " w" picW : " w" w ) . (h = "" ? " h" picH : " h" h ) " vAG" AGcount, Shell.Explorer
	AG%AGcount%.navigate("about:blank")
	AG%AGcount%.document.write(html)
	return "AG" AGcount
}

