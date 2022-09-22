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
#Include %A_ScriptDir%\Ortho_Locations.ahk
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

f1:: ; Advanced search active RXWizard patient in 3Shape using script number
{
	Gui_progressBar(action:="create")

	patientInfo := Neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

	Ortho_advSearch(patientInfo, searchMethod:="scriptNumber")
	Gui_progressBar(action:="destroy")

	return
}

f2:: ; takes snapshot from 3shape and uploads to RXWizard
{
	Gui_progressBar(action:="create")

	patientInfo := Neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=25)

	screenshotDir := Ortho_takeBitePics(patientInfo)
	Gui_progressBar(action:="update", percent:=50)

	Neo_Activate(scanField:=false)
	Gui_progressBar(action:="update", percent:=75)

	Neo_uploadPic(screenshotDir)
	Gui_progressBar(action:="destroy")

	return
}

f3:: ; just for testing, for now
{
	Gui_progressBar(action:="create")
	Ortho_export()

	Gui_progressBar(action:="destroy")
	return
}

; =========================================================================================================================
; RXWizard Shortcuts for importers
; =========================================================================================================================
#IfWinNotActive ahk_class Notepad

f4:: ; place cursor in the search field in RXWizard for barcode scanning
{
	Gui_progressBar(action:="create")

    Neo_activate(scanField:=True)
	Gui_progressBar(action:="destroy")
    return
}

f5:: ; Swap between review and edit pages
{
	Gui_progressBar(action:="create")

	Neo_swapPages(destPage:="swap")
	Gui_progressBar(action:="destroy")
	return
}

f6:: ; enters Cadent ID into RXWizard note
{
	Gui_progressBar(action:="create")

	orderID := Cadent_getOrderID()
	Gui_progressBar(action:="update", percent:=30)

	Neo_Activate(scanField:=false)
	Gui_progressBar(action:="update", percent:=60)

	Neo_NewNote(orderID)
	Gui_progressBar(action:="destroy")
	return
}

; =========================================================================================================================
; Ortho Analyzer Shortcuts
; =========================================================================================================================

f7:: ; retrieve patient info from RXWizard and perform advanced search inside OrthoAnalyzer
{
	Gui_progressBar(action:="create")

	patientInfo := Neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

	Ortho_advSearch(patientInfo, searchMethod:="patientName")
	Gui_progressBar(action:="destroy")
    return
}

f8:: ; Create new patient if non exists, then create model set
{
	Gui_progressBar(action:="create")

	patientInfo := Neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

	Ortho_createModelSet(patientInfo)
	Gui_progressBar(action:="destroy")
	return
}

f9:: ; for importing, returns to patient info and deletes temp STLs
{
	Gui_progressBar(action:="create")

	if !WinExist("OrthoAnalyzer. Patient ID:") 
	{
		MsgBox Must be in case view window in Ortho Analyzer for this function
		Exit
	}
	SetTitleMatchMode, 1
	WinActivate OrthoAnalyzer. Patient ID:
	WinWaitActive OrthoAnalyzer. Patient ID:

	quickClick(3shapeButtons["patientBrowserX"], 3shapeButtons["patientBrowserY"])
	Gui_progressBar(action:="update", percent:=50)

	if FileExist(A_MyDocuments "\Temp Models\*.stl") 
	{
		Loop, 10
			FileDelete, %A_MyDocuments%\Temp Models\*.stl
	}

	Gui_progressBar(action:="destroy")
    return
}

; =========================================================================================================================
; OrthoCAD Shortcuts
; =========================================================================================================================

f10:: ; get patient info from RXWizard and search in myCadent
{
	Gui_progressBar(action:="create")

	if !Cadent_stillOpen()
	{
		Gui_progressBar(action:="destroy")
		return
	}
	
	patientInfo := Neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=50)

    Cadent_ordersPage(patientInfo, patientSearch:=True)
	Gui_progressBar(action:="destroy")
	return
}

f11:: ; While on patient page in myCadent, export STL
{
	Gui_progressBar(action:="create")

	patientInfo := Neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=10)

	currentURL := Cadent_stillOpen()
	Gui_progressBar(action:="update", percent:=20)

	Cadent_exportClick(currentURL) ; exports through the myCadent site, opens OrthoCAD
	Gui_progressBar(action:="update", percent:=30)

	exportFilename := Cadent_exportOrthoCAD(patientInfo) ; exports the STL from OrthoCad, closes OrthoCAD
	Gui_progressBar(action:="update", percent:=70)

	Cadent_moveSTLs(exportFilename) ; moves exported STLs to the temp models folder
	Gui_progressBar(action:="destroy")
    return
}

f12:: ; renames arches in temp models folder, asks user for arch selection and auto vs manual importing, moves files
{
	Gui_progressBar(action:="create")

	patientInfo := Neo_getPatientInfo()
	Gui_progressBar(action:="update", percent:=10)

	existingArchFilenames := parseArches()
	Gui_progressBar(action:="destroy")

	filenameBase := patientInfo["engravingBarcode"] "~" patientInfo["firstName"] "~" patientInfo["lastName"] "~" patientInfo["clinicName"] "~" patientInfo["scriptNumber"] "~"
	filenameBase := StrReplace(filenameBase, " ", "_")

	finishOptions := Gui_finishImport(existingArchFilenames["arches"])

	finalizeSTLs(finishOptions, existingArchFilenames, filenameBase)
	return
}

Insert::
{
	Neo_swapPages(destPage:="cases", assignedCases:=True)
	return
}

; ===========================================================================================================================
; 3D Mouse Button Functions for use in Ortho Analyzer or Appliance Designer
; ===========================================================================================================================

SettitleMatchMode, 1
#IfWinExist ahk_group ThreeShape

!+t:: ; top/bottom view
{
    topTick := Ortho_view(3shapeButtons["topViewY"], 3shapeButtons["bottomViewY"], topTick)
	return
}

!+r:: ; right/left view
{
	sideTick := Ortho_view(3shapeButtons["rightViewY"], 3shapeButtons["leftViewY"], sideTick)
	return
}

!+f:: ; front/back view
{
	frontTick := Ortho_view(3shapeButtons["frontViewY"], 3shapeButtons["backViewY"], frontTick)
	return
}

!+c::  ; Transparency
{
	Ortho_transparency()
	return
}

!+v:: ; Switch Visible Model
{
	Ortho_visibleModel()
	return
}

!+g:: ; hits "next" button in 3shape prepping
{
	readyToExport := Ortho_nextButton()
	if (readyToExport = true)
	{
		Gui_progressBar(action:="create", percent:=50)
		Ortho_export()
		Gui_progressBar(action:="destroy")
	}
	return
}

!+1:: ; Wax Knife preset double tap tools
{
	waxOneTick := Ortho_wax(firstKnife:=1, secondKnife:=5, lastTick:=waxOneTick)
	return
}

!+2::
{
	waxTwoTick := Ortho_wax(firstKnife:=2, secondKnife:=6, lastTick:=waxTwoTick)
	return
}

!+3::
{
	waxThreeTick := Ortho_wax(firstKnife:=3, secondKnife:=7, lastTick:=waxThreeTick)
	return
}

!+5::
{
	Ortho_wax(firstKnife:=4, secondKnife:=4, lastTick:=waxTwoTick)
	return
}

!+a:: ; Artifact Removal
{
	Ortho_toolSelect(selectedTool:="artifact")
	return
}

!+p:: ; Plane Cut
{
	Ortho_toolSelect(selectedTool:="planeCut")
	return
}

!+s:: ; Spline Cut
{
	Ortho_toolSelect(selectedTool:="splineCut")
	return
}


; =========================================================================================================================
; Extra Functions
; =========================================================================================================================

quickClick(xCoord, yCoord, secondXCoord:=false, secondYCoord:=false, pauseTime:=0)
{
	BlockInput MouseMove
    MouseGetPos, x, y
    Click, %xCoord%, %yCoord%
	if (secondXCoord != false)
	{
		Sleep, %pauseTime%
		Click, %secondXCoord%, %secondYCoord%
	}
    MouseMove, %x%, %y%, 0
    BlockInput MouseMoveOff
	return
}

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

	tagsUpper := ["Upr.stl", "~Upr", "u.stl", "upper.stl", "max.stl", "Max.stl", "_u-"]
	tagsLower := ["Lwr.stl", "~Lwr", "l.stl", "lower.stl", "man.stl", "Man.stl", "_l-"]

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
    if (finishOptions["arches"] = True)
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
		AddAnimatedGIF(AnimPicFile)
		Gui, Show
		return
	}
}

AddAnimatedGIF(imagefullpath , x="", y="", w="", h="", guiname = "1")
{
	global AG1
	static pic
	html := "<html><body style='background-color: transparent' style='overflow:hidden' leftmargin='0' topmargin='0'><img src='" imagefullpath "' width=" w " height=" h " border=0 padding=0></body></html>"
	Gui, AnimGifxx:Add, Picture, vpic, %imagefullpath%
	GuiControlGet, pic, AnimGifxx:Pos
	Gui, AnimGifxx:Destroy
	Gui, %guiname%:Add, ActiveX, % (x = "" ? " " : " x" x ) . (y = "" ? " " : " y" y ) . (w = "" ? " w" picW : " w" w ) . (h = "" ? " h" picH : " h" h ) " vAG1", Shell.Explorer
	AG1.navigate("about:blank")
	AG1.document.write(html)
	return
}

#Include %A_ScriptDir%\Libraries\Gui.ahk