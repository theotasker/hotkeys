; ===========================================================================================================================
; Combined functions for Importing, Prepping, and Engraving, using netfabb, 3Shape, and the websites
; ===========================================================================================================================

#SingleInstance, Force

; Library includes all Ortho_ functions
#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\OrthoLibrary.ahk

; Library includes all Netfabb_ functions
#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\NetfabbLibrary.ahk

; Library includes all Web_,  Neo_, and Cadent_ functions
#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\WebLibrary.ahk

; Library ONLY for CSS and XPATH locations of items in web functions
#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\WebPathLibrary.ahk

; Library specifically to capture screenshots, must be in 32bit mode in AHK
#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\CaptureScreen.ahk

; Library for logging total hotkeys, clicks saved, and cases complete (Log_ functions)
#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\LogLibrary.ahk

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
	StartStep := "Complete Digital Import"
	EndStep := "End Digital Import"
	Gui, Destroy
	return
}

Prepping:
{
	Step := "Prepping"
	StartStep := "Complete Digital Prep"
	EndStep := "End Digital Prep"
	Gui, Destroy
	Exit
}

Engraving:
{
	Step := "Engraving"
	StartStep := "Complete Engraving"
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
    Neo_NavigateReviewFromEdit()
	Log_Increment(2, 0)
	Gui, Destroy
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
		
		
	if managercheck = no ; checks for the manager.txt doc, will click the generate schedule button if not
		
		; Click the "generate schedule" button
		try NeoDriver.findElementByCss(Path_GenerateScheduleCSS).Click()
		catch e 
		{
			MsgBox,, Couldn't Find Element, Couldn't find the "Generate Schedule" button
			Gui, Destroy
			Exit
		}
	
        
	GuiControl,, Progress, 33
    
	; Current step variables
	currentstep := StartStep
	
	
	if managercheck = no
	{
		; xpath for non-managers
		currentstepxpath := Path_StartStopXPATH
	}
	if managercheck = yes
	{
		; xpath for managers
		currentstepxpath := Path_StartStopManagerXPATH
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
	Neo_GetInfoFromEdit()
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
        ControlFocus, TEdit4
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
	
	Neo_GetInfoFromEdit()
	
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
    global firstname, lastname, ProgX, ProgY
    
    Gui, Add, Progress, vprogress w300 h45
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop
    
    ; title must match exactly
    SetTitleMatchMode, 3
    if !WinExist("OrthoCAD Export ahk_exe OrthoCAD.exe")
    {
        MsgBox This shortcut requires the "OrthoCAD Export" window to be open
        Gui, Destroy
        Exit
    }
    
    BlockInput, MouseMove
    
    GuiControl,, Progress, 20
    
    ; Activate orthocad, then the export window
    WinActivate, ahk_exe OrthoCAD.exe
    WinActivate, "OrthoCAD Export"
    
    GuiControl,, Progress, 40
    
    ; set the export type to open scan
    ControlFocus, ComboBox1, OrthoCAD Export
    Send, {down}
    Sleep, 100
    
    GuiControl,, Progress, 60
    
    ; set the export type to models in occlusion
    ControlFocus, ComboBox2, OrthoCAD Export
    Send, {down}{down}
    Sleep, 100
    
    GuiControl,, Progress, 80
    
    ; set the export folder to "export"
    ControlFocus, Edit1, OrthoCAD Export
    Send {CtrlDown}a{CtrlUp}
    Sleep, 100
    Send, %firstname%%lastname%
    Sleep, 100
    
    GuiControl,, Progress, 90
    
    ; go to the export button and hit enter
    Send, {tab}
    Sleep, 100
    Send, {enter}
    
    GuiControl,, Progress, 100
    Gui, Destroy
    
    
    BlockInput, MouseMoveOff
	
	Log_Increment(8, 0)
	
    return
}
    
; Finish the orthocad export    
f12::
{
    ; Move STLs in the named export folder
	global firstname, lastname, ProgX, ProgY
    
    Gui, Add, Progress, vprogress w300 h45
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop
    
    newfirst := StrReplace(firstname, " ", "")
    newlast := StrReplace(lastname, " ", "")
	
	IfNotExist, C:\Cadent\Export\
		FileCreateDir, C:\Cadent\Export\
	
	IfNotExist, %A_MyDocuments%\Temp Models
		FileCreateDir, %A_MyDocuments%\Temp Models
	
	IfNotExist, C:\Cadent\Export\%newfirst%%newlast%
		MsgBox Couldn't find export folder, didn't move files
    
    GuiControl,, Progress, 20
    
    BlockInput, MouseMove
	
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
    
    GuiControl,, Progress, 40
    
    ; Close OrthoCad if it's open
    SetTitleMatchMode, 3
    if WinExist("Export Done ahk_exe OrthoCAD.exe")
    {	

        WinActivate, Export Done
        WinWaitActive, Export Done,, 20
        if ErrorLevel
        {
            BlockInput, MouseMoveOff
            MsgBox,, Window Activation Error, Couldn't activate Orthocad
            return
        }
    
        Sleep, 100
        ControlFocus, Button1, Export Done
        Sleep, 300
        Send, {tab}
        Sleep, 100
        Send, {enter}
    }
    
    GuiControl,, Progress, 60
    
    WinClose, ahk_class OrthoCAD
    Sleep, 200
    
    WinActivate, ahk_exe OrthoCAD.exe
    
    Sleep, 100
    Send {tab}
    Sleep, 100
    Send {Enter}
    
    BlockInput, MouseMoveOff
    
    GuiControl,, Progress, 80
    
    Cadent_Orders()
    
    GuiControl,, Progress, 100
    Gui, Destroy
	
	Log_Increment(6, 2)
    
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
		Neo_GetInfoFromEdit()
		
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

		
		WinWaitActive Open,, 10
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

f3::
{
	Netfabb_Loop()
	Log_Increment(3, 2)
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