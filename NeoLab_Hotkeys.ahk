; ===========================================================================================================================
; Combined functions for Importing, Prepping, 3Shape, and the websites
; ===========================================================================================================================

; ===========================================================================================================================
; AutoHotKey Startup Stuff
; ===========================================================================================================================

#SingleInstance, Force 
CoordMode, Mouse, Client ; mouse and pixel coordinates will be based on the client, instead of screen or window. Most precise
CoordMode, Pixel, Client

#Include %A_ScriptDir%\Libraries\CaptureScreen.ahk ; 3rd party library, must be in 32bit mode in AHK
#Include %A_ScriptDir%\Web_Paths.ahk
#Include %A_ScriptDir%\Libraries\Neo_Functions.ahk
#Include %A_ScriptDir%\Libraries\Cadent_Functions.ahk
#Include %A_ScriptDir%\Libraries\Ortho_Functions.ahk

; default progress bar location
global progressBarX := "x788"
global progressBarY := "y150"

; ===========================================================================================================================
; Directories
; ===========================================================================================================================

global autoImportDir := "\\NEO-AUTOMATE\AutoImport\Input\"
global tempModelsDir := A_MyDocuments "\Temp Models\"
global screenshotDir := A_MyDocuments "\Automation\Screenshots"

; ===========================================================================================================================
; Prepping Shortcuts
; ===========================================================================================================================
#IfWinNotActive ahk_exe explorer.exe ; these shouldn't overwrite the default windows functions

f1::
{
	Gui_progressBar(action:="create", percent:=0)

	patientInfo := neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

	Ortho_AdvSearch(patientInfo, searchMethod:="scriptNumber")
	Gui_progressBar(action:="destroy", percent:=0)

	return
}

f2:: 
{
	Gui_progressBar(action:="create", percent:=0)

	patientInfo := neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=25)

	screenshotDir := Ortho_takeBitePics(patientInfo)
	Gui_progressBar(action:="update", percent:=50)

	Neo_Activate(scanField:=false)
	Gui_progressBar(action:="update", percent:=75)

	Neo_uploadPic(screenshotDir)
	Gui_progressBar(action:="destroy", percent:=0)

	return
}

f3:: ; just for testing, for now
{
	Gui_progressBar(action:="create", percent:=0)
	Ortho_Export()

	Gui_progressBar(action:="destroy", percent:=0)
	return
}

; =========================================================================================================================
; RXWizard Shortcuts for importers
; =========================================================================================================================
#IfWinNotActive ahk_class Notepad

f4:: ; place cursor in the search field in RXWizard for barcode scanning
{
	Gui_progressBar(action:="create", percent:=0)

    Neo_activate(scanField:=True)
	Gui_progressBar(action:="destroy", percent:=0)
    return
}

f5:: ; Swap between review and edit pages
{
	Gui_progressBar(action:="create", percent:=0)

	Neo_swapPages(destPage:="swap")
	Gui_progressBar(action:="destroy", percent:=0)
	return
}

f6:: ; hit start/stop button for case, must be on review or edit page
{
	Gui_progressBar(action:="create", percent:=0)

	orderID := Cadent_GetOrderID()
	Gui_progressBar(action:="update", percent:=30)

	Neo_Activate(scanField:=false)
	Gui_progressBar(action:="update", percent:=60)

	Neo_NewNote(orderID)
	Gui_progressBar(action:="destroy", percent:=0)
	return
}

; =========================================================================================================================
; Ortho Analyzer Shortcuts
; =========================================================================================================================

f7:: ; retrieve patient info from RXWizard and perform advanced search inside OrthoAnalyzer
{
	Gui_progressBar(action:="create", percent:=0)

	patientInfo := neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

	Ortho_AdvSearch(patientInfo, searchMethod:="patientName")
	Gui_progressBar(action:="destroy", percent:=0)
    return
}

f8:: ; Create new patient if non exists, then create model set
{
	Gui_progressBar(action:="create", percent:=0)

	patientInfo := neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

	ortho_createModelSet(patientInfo)
	Gui_progressBar(action:="destroy", percent:=0)
	return
}

f9:: ; for importing, returns to patient info and deletes temp STLs
{
	Gui_progressBar(action:="create", percent:=0)

	if !WinExist("OrthoAnalyzer. Patient ID:") 
	{
		MsgBox Must be in case view window in Ortho Analyzer for this function
		Exit
	}
	SetTitleMatchMode, 1
	WinActivate OrthoAnalyzer. Patient ID:
	WinWaitActive OrthoAnalyzer. Patient ID:

	quickClick("27", "43")
	Gui_progressBar(action:="update", percent:=50)

	if FileExist(A_MyDocuments "\Temp Models\*.stl") 
	{
		Loop, 10
			FileDelete, %A_MyDocuments%\Temp Models\*.stl
	}

	Gui_progressBar(action:="destroy", percent:=0)
    return
}

; =========================================================================================================================
; OrthoCAD Shortcuts
; =========================================================================================================================

f10:: ; get patient info from RXWizard and search in myCadent
{
	Gui_progressBar(action:="create", percent:=0)

	if !Cadent_StillOpen()
	{
		Gui_progressBar(action:="destroy", percent:=0)
		return
	}
	
	patientInfo := neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

    Cadent_ordersPage(patientInfo, patientSearch:=True)
	Gui_progressBar(action:="destroy", percent:=0)
	return
}

f11:: ; While on patient page in myCadent, export STL
{
	Gui_progressBar(action:="create", percent:=0)

	patientInfo := neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=10)

	currentURL := Cadent_StillOpen()
	Gui_progressBar(action:="update", percent:=20)

	Cadent_exportClick(currentURL) ; exports through the myCadent site, opens OrthoCAD
	Gui_progressBar(action:="update", percent:=30)

	exportFilename := Cadent_exportOrthoCAD(patientInfo) ; exports the STL from OrthoCad, closes OrthoCAD
	Gui_progressBar(action:="update", percent:=70)

	Cadent_moveSTLs(exportFilename) ; moves exported STLs to the temp models folder
	Gui_progressBar(action:="destroy", percent:=0)
    return
}

f12:: ; renames arches in temp models folder, asks user for arch selection and auto vs manual importing, moves files
{
	Gui_progressBar(action:="create", percent:=0)

	patientInfo := neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=10)

	existingArchFilenames := parseArches()
	Gui_progressBar(action:="destroy", percent:=20)

	filenameBase := patientInfo["engravingBarcode"] "~" patientInfo["firstName"] "~" patientInfo["lastName"] "~" patientInfo["clinicName"] "~" patientInfo["scriptNumber"] "~"
	filenameBase := StrReplace(filenameBase, " ", "_")

	finishOptions := Gui_finishImport(existingArchFilenames["arches"])

	finalizeSTLs(finishOptions, existingArchFilenames, filenameBase)
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

!+c::  ; Transparency
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
	{
		return
	}

	WinActivate ahk_group ThreeShape
	quickClick(artifactX, artifactY)
	return
}

; Plane Cut
!+p::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular 
	{
		return
	}

	WinActivate ahk_group ThreeShape
	quickClick(planeCutX, planeCutY)
	return
}

; Spline Cut
!+s::
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular 
	{
		return
	}

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

; =========================================================================================================================
; Extra Functions
; =========================================================================================================================

parseArches() {
    if !FileExist(tempModelsDir "*.stl")
	{
		MsgBox,, Importing Finishing Error, No STLs in the Temp Models Folder
		return
	}

	tempModelList := []  ; list of all STLs in temp models folder
	Loop Files, %tempModelsDir%*.stl 
		tempModelList.Push(A_LoopFileName)

	if(tempModelList.Length() > 2)
	{
		MsgBox,, Importing finishing Error, Too many STLs in Temp Models folder
		return
	}

	tagsUpper := ["Upr.stl", "u.stl", "upper.stl", "max.stl", "Max.stl", "_u-"]
	tagsLower := ["Lwr.stl", "l.stl", "lower.stl", "man.stl", "Man.stl", "_l-"]

    archFilenames := {"arches":False, "upper":False, "lower":False}
	tagCheck := False
	for key, filename in tempModelList  ; for each stl in temp folder
	{
		for key, tag in tagsUpper ; check against upper tags
		{
			if instr(filename, tag) ; if the filename has this upper tag
			{
				if (archFilenames["upper"] != False)  ; if there's already an upper saved
				{
					msgbox,, Importing Finishing Error, found more than one upper STL in Temp Models folder
					return
				}
				archFilenames["upper"] := filename ; otherwise, assign it to the upper
				tagCheck := True
			}
		}
		for key, tag in tagsLower
		{
			if instr(filename, tag)
			{
				if (archFilenames["lower"] != False)
				{
					msgbox,, Importing Finishing Error, found more than one lower STL in Temp Models folder
					return
				}
				archFilenames["lower"] := filename
				tagCheck := True
			}
		}
	}

	if (tagCheck := False)
	{
		msgbox,, Importing Finishing Error, STLs must be labeled as upper or lower
        return
	}

    if ((archFilenames["upper"] != False) and (archFilenames["lower"] != False))
    {
        archFilenames["arches"] := "both"
    }
    else if (archFilenames["upper"] != False)
    {
        archFilenames["arches"] := "upper"
    }
    else if (archFilenames["lower"] != False)
    {
        archFilenames["arches"] := "lower"
    }

    return archFilenames
}

finalizeSTLs(finishOptions, existingArchFilenames, filenameBase) {
    if (finishOptions["arches"] = "both")
    {
        filenameTag := "[2].stl"
    }
    Else
    {
        filenameTag := "[1].stl"
    }

    if (finishOptions["auto"] = True)
    {
        destinationDir := autoImportDir
    }
    Else
    {
        destinationDir := tempModelsDir
    }

    if (finishOptions["upper"] = True)
    {
        currentFullFilename := tempModelsDir existingArchFilenames["upper"]
        destFullFilename := destinationDir filenameBase "Upr" filenameTag
        FileMove, %currentFullFilename%, %destFullFilename%
    }
    if (finishOptions["lower"] = True)
    {
        currentFullFilename := tempModelsDir existingArchFilenames["lower"]
        destFullFilename := destinationDir filenameBase "Lwr" filenameTag
        FileMove, %currentFullFilename%, %destFullFilename%
    }
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
		AnimPicFile := A_ScriptDir "\Files\JP.gif"
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

#Include %A_ScriptDir%\Libraries\Gui.ahk