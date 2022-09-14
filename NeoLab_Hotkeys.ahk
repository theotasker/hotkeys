; ===========================================================================================================================
; Combined functions for Importing, Prepping, and Engraving, using netfabb, 3Shape, and the websites
; ===========================================================================================================================

; ===========================================================================================================================
; AutoHotKey Startup Stuff
; ===========================================================================================================================

#SingleInstance, Force 
CoordMode, Mouse, Client ; mouse and pixel coordinates will be based on the client, instead of screen or window. Most precise
CoordMode, Pixel, Client

#Include D:\hotkeys\Libraries\CaptureScreen.ahk ; 3rd party library, must be in 32bit mode in AHK
#Include D:\hotkeys\Web_Paths.ahk
#Include D:\hotkeys\Libraries\Neo_Functions.ahk
#Include D:\hotkeys\Libraries\Cadent_Functions.ahk
#Include D:\hotkeys\Libraries\Netfabb_Functions.ahk
#Include D:\hotkeys\Libraries\Ortho_Functions.ahk

SetBox() ; opens the settings box to select step

; ===========================================================================================================================
; Engraving shortcuts in Netfabb
; ===========================================================================================================================

#IfWinNotActive ahk_exe explorer.exe ; these shouldn't overwrite the default windows functions

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

; =========================================================================================================================
; RXWizard Shortcuts
; =========================================================================================================================
#IfWinNotActive ahk_class Notepad

f4:: ; place cursor in the search field in RXWizard for barcode scanning
{
    neo_activate(scanField:=True)
    return
}

f5:: ; Swap between review and edit pages
{
	neo_swapPages(destPage:="swap")
	return
}

f6:: ; hit start/stop button for case, must be on review or edit page
{
    neo_swapPages(destPage:="review")

	needStop := neo_start(currentStep:=currentStep)

	if (Needstop = True)
	{
		neo_stop(currentStep:=currentStep)
	}
	Sleep, 1000

	neo_activate(scanField:=true)

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

	return
}

f9:: ; for importing, returns to patient info and deletes temp STLs. for prepping, returns to patient info and exports STL 
{
	if currentStep = "importing" {
		if !WinExist("OrthoAnalyzer. Patient ID:") {
			MsgBox Must be in case view window in Ortho Analyzer for this function
			Exit
		}
		SetTitleMatchMode, 1
		WinActivate OrthoAnalyzer. Patient ID:
		WinWaitActive OrthoAnalyzer. Patient ID:

		quickClick("27", "43")

		if FileExist(A_MyDocuments "\Temp Models\*.stl") {
			Loop, 10
				FileDelete, %A_MyDocuments%\Temp Models\*.stl
		}
	}

	else if currentStep = "prepping"
	{
		if WinExist("ahk_group ThreeShape") {
			Ortho_Export()
		}
		else {
			MsgBox,, Wrong Window, Must be in a case to use this function
			Exit
		}
	}

	else {
		MsgBox,, Wrong Step, Must be on the importing or prepping step for this function
	}

    return
}

; =========================================================================================================================
; OrthoCAD Shortcuts
; =========================================================================================================================

f10:: ; get patient info from RXWizard and search in myCadent
{
	patientInfo := neo_getInfoFromReview()

    Cadent_ordersPage(patientInfo, patientSearch:=True)

	return
}

f11:: ; While on patient page in myCadent, export STL
{
	patientInfo := neo_getInfoFromReview()

	currentURL := Cadent_StillOpen()

	Cadent_exportClick()

	exportFilename := Cadent_exportOrthoCAD(patientInfo)

	Cadent_moveSTLs(exportFilename)

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

	if Step = Importing ; On the importing step, grabs the iTero ID and inserts it into a new note
	{
		orderID := Cadent_GetOrderID()

		Neo_Activate(scanField:=false)

		Neo_NewNote(orderID)

		return
	}

	else ; If on any other step, takes bite screenshots in ortho and uploads them to the website
	{
		patientInfo := neo_getInfoFromReview()

		Ortho_takeBitePics(patientInfo)

		Neo_Activate(scanField:=false)

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