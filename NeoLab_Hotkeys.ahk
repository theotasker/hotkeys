; ===========================================================================================================================
; Combined functions for Importing, Prepping, and Engraving, using netfabb, 3Shape, and the websites
; ===========================================================================================================================

; version test

; ===========================================================================================================================
; AutoHotKey Startup Stuff
; ===========================================================================================================================

#SingleInstance, Force

CoordMode, Mouse, Client ; mouse and pixel coordinates will be based on the client, instead of screen or window. Most precise
CoordMode, Pixel, Client

#Include \\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Projects\CaptureScreen.ahk ; 3rd party library, must be in 32bit mode in AHK
#Include D:\hotkeys\Web_Paths.ahk
#Include D:\hotkeys\Libraries\Neo_Functions.ahk
#Include D:\hotkeys\Libraries\Cadent_Functions.ahk
#Include D:\hotkeys\Libraries\Netfabb_Functions.ahk
#Include D:\hotkeys\Libraries\Ortho_Functions.ahk

; starter variable for the F2 function in netfabb
placedengraving =: 0

; Variables set at the start to see if any webdrivers have been started previously
NeoCheck := 0
Cadentcheck := 0

SetBox() ; opens the settings box to select step

; ===========================================================================================================================
; Engraving shortcuts in Netfabb
; ===========================================================================================================================

; starter variable for the F2 function in netfabb
placedengraving =: 0

#IfWinNotActive ahk_exe explorer.exe

f1::
{
	Netfabb_Level()
	return
}

f2::
{
	Netfabb_finish()
	return
}


#IfWinNotActive ahk_class Notepad

; =========================================================================================================================
; RXWizard Shortcuts
; =========================================================================================================================

; shortcut to call the home function and return to the cases page
f4::
{
    neo_navigateToCases()
    return
}

; shortcut to navigate to review from edit page, must be on edit page
f5::
{

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

	msgbox % webstep

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


	GuiControl,, Progress, 33

	; Current step variables
	currentstep := StartStep



	try NeoDriver.executeScript("window.scrollBy(0, 800)")
	catch e
	{
		msgbox couldn't scroll
		return
	}

	; calls Neo_ function to click the start button
	neo_start()

	GuiControl,, Progress, 66

	currentstep := EndStep

	; calls Neo_ function to click the stop button
	neo_stop()

	Sleep, 1000

	GuiControl,, Progress, 100
	Gui, Destroy

	neo_navigateToCases()


	return
}

; =========================================================================================================================
; Ortho Analyzer Shortcuts
; =========================================================================================================================

f7:: ; retrieve patient info from RXWizard and perform advanced search inside OrthoAnalyzer
{
	patientInfo := neo_getInfoFromReview()

	Ortho_AdvSearch(patientInfo)

	BlockInput, MouseMoveOff
    return
}

f8:: ; Create new patient if non exists, then create model set
{
	patientInfo := neo_getInfoFromReview()

	ortho_createModelSet(patientInfo)

	BlockInput, MouseMoveOff
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

		quickClick("27", "43")

		if FileExist(A_MyDocuments "\Temp Models\*.stl")
			Loop, 10
				FileDelete, %A_MyDocuments%\Temp Models\*.stl
	}

	else if Step = Prepping
	{
		if WinExist("ahk_group ThreeShape")
		{
			Ortho_Export()
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

	patientInfo := neo_getInfoFromReview()

    ;Progress Bar instance
    Gui, Add, Progress, vprogress w300 h45
	Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
    Gui, +AlwaysOnTop

    GuiControl,, Progress, 33

    Cadent_Orders()

    GuiControl,, Progress, 66

    BlockInput, MouseMove

    if StrLen(patientInfo["firstName"]) > 1 and StrLen(patientInfo["lastName"]) > 1
	{
        Send % patientInfo["lastName"]
		Send ,{space}
		Send % patientInfo["firstName"]
	}

    BlockInput, MouseMoveOff

    GuiControl,, Progress, 100
    Gui, Destroy

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

	neo_getInfoFromReview()

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

	if Step = Importing ; On the importing step, grabs the iTero ID and inserts it into the notes field
	{
		Gui, Add, Progress, vprogress w300 h45
		Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
		Gui, +AlwaysOnTop

		orderID := Cadent_GetOrderID()

		GuiControl,, Progress, 10

		Neo_Activate()

		GuiControl,, Progress, 20

		Neo_NewNote(orderID)

		GuiControl,, Progress, 75

		;tab to the save button, css path doesn't work anymore
		Send {tab}{tab}

		Sleep, 200

		Send {enter}

		GuiControl,, Progress, 100
		Gui, Destroy

	}

	else ; If on any other step, takes bite screenshots in ortho and uploads them to the website
	{

		; gets first and last name to be used in the screenshot filenames
		neo_stillOpen()
		patientInfo := neo_getInfoFromReview()

		if !WinExist("ahk_group ThreeShape")
		{
			MsgBox,, Wrong Window, Case must be open in OrthoAnalyzer or ApplianceDesigner
			Exit
		}

		; defines directory and destroys and recreates it to ensure it's empty
		dir= %A_MyDocuments%\Automation\Screenshots
		FileRemoveDir, %dir%, 1
		FileCreateDir, %dir%

		Ortho_View(frontViewY, bottomViewY, topTick)

		screenshotname := patientInfo["firstName"] patientInfo["lastName"] "Front.jpg"
		Sleep, 500
		CaptureScreen(screenshotname) ; function defined in CaptureScreen Library

		Ortho_View(leftViewY, bottomViewY, topTick)

		screenshotname := patientInfo["firstName"] patientInfo["lastName"] "Left.jpg"
		Sleep, 500
		CaptureScreen(screenshotname)

		Ortho_View(rightViewY, bottomViewY, topTick)

		screenshotname := patientInfo["firstName"] patientInfo["lastName"] "Right.jpg"
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


    return
}

; ===========================================================================================================================
; 3D Mouse Button Functions for use in Ortho Analyzer or Appliance Designer
; ===========================================================================================================================

SettitleMatchMode, 1
#IfWinExist ahk_group ThreeShape

!+t:: ; top/bottom view
{
    topTick := Ortho_View(topViewY, bottomViewY, topTick)
	return
}

!+r:: ; right/left view
{
	sideTick := Ortho_View(rightViewY, leftViewY, sideTick)
	return
}

!+f:: ; front/back view
{
	frontTick := Ortho_View(frontViewY, backViewY, frontTick)
	return
}

!+c:: ; Transparency
{
	Ortho_View(transparencyY, transparencyY, topTick)
	return
}

!+v:: ; Switch Visible Model
{
	UpperDotX := 1708
	UpperDotY := 79
	LowerDotX := 1708
	LowerDotY := 106

	Ortho_VisibleModel()
	return
}

!+.:: ; Export finished case
{
	Ortho_Export()
	return
}

!+g:: ; hits "next" button in 3shape prepping
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape

	if InStr("Prepare occlusion,Setup plane alignment,Virtual base,Sculpt maxillary,Sculpt mandibular", PrepStep)
		WinActivate ahk_group ThreeShape
		quickClick(nextButtonX, nextButtonY)
	return
}

; Wax Knife preset double tap tools
!+1::
{
	waxOneTick := Ortho_Wax(firstKnife:=1, secondKnife:=5, lastTick:=waxOneTick)
	return
}

!+2::
{
	waxTwoTick := Ortho_Wax(firstKnife:=2, secondKnife:=6, lastTick:=waxTwoTick)
	return
}

!+3::
{
	waxThreeTick := Ortho_Wax(firstKnife:=3, secondKnife:=7, lastTick:=waxThreeTick)
	return
}

!+5::
{
	Ortho_Wax(firstKnife:=4, secondKnife:=4, lastTick:=waxTwoTick)
	return
}

; Artifact Removal
!+a::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
		return

	WinActivate ahk_group ThreeShape
	quickClick(artifactX, artifactY)
	return
}

; Plane Cut
!+p::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
		return

	WinActivate ahk_group ThreeShape
	quickClick(planeCutX, planeCutY)
	return
}

; Spline Cut
!+s::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
		return

	WinActivate ahk_group ThreeShape
	quickClick(splineCutX, splineCutY)
	BlockInput MouseMove
	Sleep, 100
	WinActivate ahk_group ThreeShape
	quickClick(splineSmoothX, splineSmoothY)
	return
}


quickClick(xCoord, yCoord)
{
	BlockInput MouseMove
    MouseGetPos, x, y
    Click, %xCoord%, %yCoord%
    MouseMove, %x%, %y%, 0
    BlockInput MouseMoveOff
	return
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

#Include D:\hotkeys\Libraries\Gui.ahk