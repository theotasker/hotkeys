; ===========================================================================================================================
; Library for Ortho_ functions
; ===========================================================================================================================

; ===========================================================================================================================
; Variables and groups
; ===========================================================================================================================

; starting ticks for double-tap functions for 3D mouse
topTick := A_TickCount
sideTick := A_TickCount
frontTick := A_TickCount

waxOneTick := A_TickCount
waxTwoTick := A_TickCount
waxThreeTick := A_TickCount

; Groups to determine active windows
GroupAdd, ThreeShape, OrthoAnalyzer
GroupAdd, ThreeShape, ApplianceDesigner

GroupAdd, ThreeShapePatient, New patient model set info
GroupAdd, ThreeShapePatient, New patient info

GroupAdd, ThreeShapeExe, ahk_exe OrthoAnalyzer.exe
GroupAdd, ThreeShapeExe, ahk_exe ApplianceDesigner.exe

; ===========================================================================================================================
; Data entry functions
; ===========================================================================================================================

Ortho_AdvSearch(patientInfo, searchMethod) ; function to enter patient name into advanced search field and search using globals
{
	BlockInput MouseMove

    if StrLen(patientInfo["firstName"]) < 2 or StrLen(patientInfo["lastName"]) < 2
    {
        BlockInput MouseMoveOff
		Gui, Destroy
        MsgBox No patient info saved from RXWizard
        Exit
    }

    if !WinExist("Open patient case")
    {
        BlockInput MouseMoveOff
		Gui, Destroy
        MsgBox Must be in the "Open Patient Case" Dialogue for this hotkey
        Exit
    }

    WinActivate, ahk_group ThreeShapeExe
    WinWaitActive, Open patient case,, 10

    if ErrorLevel
    {
		BlockInput MouseMoveOff
		Gui, Destroy
        MsgBox, Couldn't get focus on 3Shape
        return
    }
    Sleep, 200

    ControlFocus, Advanced search, Open patient case
    Sleep, 400
    Send, {Enter}
    Sleep, 400

	; ===========================================================================================================================
	; Added to check the search field locator between 2019 and 2021

	if (3shapeFields["advSearchScript"] = "")
	{
		quickClick(3shapeFields["advSearchScriptX"], 3shapeFields["advSearchScriptY"])
		Sleep, 200
		controlGetFocus, scriptHandle, Open patient case
		if (scriptHandle = 3shapeFields["advSearchScript2019"])
		{
			global 3ShapeVersion := "2019"
			3shapeFields["advSearchScript"] := 3shapeFields["advSearchScript2019"]
		}
		else if (scriptHandle = 3shapeFields["advSearchScript2021"])
		{
			global 3ShapeVersion := "2021"
			3shapeFields["advSearchScript"] := 3shapeFields["advSearchScript2021"]
		}
	}

	; ===========================================================================================================================

	if (searchMethod = "patientName")
	{
		ortho_sendText(patientInfo["firstName"], 3shapeFields["advSearchFirst"], "Open patient case")
		ortho_sendText(patientInfo["lastName"], 3shapeFields["advSearchLast"], "Open patient case")
		ortho_sendText(patientInfo["clinicName"], 3shapeFields["advSearchClinic"], "Open patient case")

		ortho_sendText("", 3shapeFields["advSearchScript"], "Open patient case")
	}
	else 
	{
		ortho_sendText(patientInfo["scriptNumber"], 3shapeFields["advSearchScript"], "Open patient case")

		ortho_sendText("", 3shapeFields["advSearchFirst"], "Open patient case")
		ortho_sendText("", 3shapeFields["advSearchLast"], "Open patient case")
		ortho_sendText("", 3shapeFields["advSearchClinic"], "Open patient case")
	}

	ControlFocus, 3shapeFields["advSearchGo"], Open patient case
	Send {enter}
	Sleep, 100
	Send {down}

    BlockInput MouseMoveOff
    return
}

ortho_createModelSet(patientInfo)
{
	BlockInput MouseMove

    if !WinExist("Open patient case")
    {
		BlockInput MouseMoveOff
		Gui, Destroy
        MsgBox must be on "Open patient case" to use this function
        Exit
    }

    WinActivate Open patient case
	quickClick(3shapeButtons["newPatientModelX"], 3shapeButtons["newPatientModelY"])

    ; see which window opens
    WinWaitActive, ahk_group ThreeShapePatient,, 5
	if ErrorLevel
    {
		BlockInput, MouseMoveOff
        Gui, Destroy
        MsgBox, Couldn't get focus patient info popup
        Exit
    }

    SetTitleMatchMode, 3
    if WinActive("New patient info",, model)
    {
		ortho_sendText(patientInfo["firstName"] patientInfo["lastName"], 3shapeFields["newPatientExt"], "New patient info")
		ortho_sendText(patientInfo["firstName"], 3shapeFields["newPatientFirst"], "New patient info")
		ortho_sendText(patientInfo["lastName"], 3shapeFields["newPatientLast"], "New patient info")
		ortho_sendText(patientInfo["clinicName"], 3shapeFields["newPatientClinic"], "New patient info")
    }
    else if WinActive("New patient model set info")
    {
		ortho_sendText(patientInfo["scriptNumber"], 3shapeFields["newModelScript"], "New patient model set info")
    }
    else
    {
        BlockInput, MouseMoveOff
		Gui, Destroy
        MsgBox,, Wrong Window, Didn't land on the "New patient info" or "New model set" window
        Exit
    }

    BlockInput, MouseMoveOff
    return
}

Ortho_Export() ; At end of model edit, waits for patient browser then exports STL
{
	BlockInput MouseMove

	WinWaitActive, Open patient case,, 30
	if ErrorLevel
	{
		BlockInput MouseMoveOff
        Gui, Destroy
		MsgBox,, Timeout Error, "Open Patient Case" window took too long to open
		Exit
	}

	Sleep, 200
	quickClick(3shapeButtons["modelExportX"], 3shapeButtons["modelExportY"])
	Sleep, 100
	Send {Up 2}
	Sleep, 100
	Send {Enter}

	WinWaitActive, Exported items,, 15 ; Wait for export confirmation box
	if ErrorLevel
	{
		BlockInput MouseMoveOff
        Gui, Destroy
		MsgBox,, Timeout Error, Export Confirmation took too long
		Exit
	}

	Send, {Enter}
	BlockInput MouseMoveOff
	return
}

ortho_sendText(textToSend, targetBox, targetWindow)
{
	ControlGetText, returnText, %targetBox%, %targetWindow%

	While(returnText != textToSend && loopCheck < 5)
	{
		ControlFocus, %targetBox%, %targetWindow%
		Sleep, 100
		Send, ^a
		Sleep, 100
		Send, {del}
		Sleep, 100
		Send, %textToSend%
		Sleep, 100
		Send, {del}
		Sleep, 100
		ControlGetText, returnText, %targetBox%, %targetWindow%
		loopCheck++
	}
	If(returnText != textToSend)
	{
		return False
	}
	else
	{
		return True
	}
}

; ===========================================================================================================================
; Prepping functions
; ===========================================================================================================================

Ortho_View(firstViewY, secondViewY, lastActionTick) ; clicks the view button declared before the function is called, swaps to a secondary view if pressed twice
{
    global toggle
	BlockInput MouseMove

    currentTick := A_TickCount
	if currentTick - lastActionTick > 1000
	{
		WinActivate ahk_group ThreeShape
		quickClick(3shapeButtons["allViewX"], firstViewY)
		toggle := 1
	}
	else
	{
		if toggle = 1
		{
			WinActivate ahk_group ThreeShape
			quickClick(3shapeButtons["allViewX"], secondViewY)

			toggle := 2
		}
		else if toggle = 2
		{
			WinActivate ahk_group ThreeShape
			quickClick(3shapeButtons["allViewX"], firstViewY)
			toggle := 1
		}
	}
	BlockInput MouseMoveOff
	return currentTick
}

Ortho_VisibleModel() ; Swaps visible Model
{
	BlockInput MouseMove

	WinActivate ahk_group ThreeShape
	PixelGetColor, upperModelPixel, 3shapeModelSlider["upperX"], 3shapeModelSlider["upperY"]
	PixelGetColor, lowerModelPixel, 3shapeModelSlider["lowerX"], 3shapeModelSlider["lowerY"]

	if ((upperModelPixel = 0x000000) and (lowerModelPixel = 0x000000)) ; both models on, turn off lower
	{
		quickClick(3shapeModelSlider["lowerX"], 3shapeModelSlider["lowerY"])
	}
	else if ((upperModelPixel = 0x000000) and (lowerModelPixel = 0xF0F0F0)) ; only upper on, swap models
	{
		quickClick(3shapeModelSlider["upperX"], 3shapeModelSlider["upperY"], 3shapeModelSlider["lowerX"], 3shapeModelSlider["lowerY"])
	}
	else if ((upperModelPixel = 0xF0F0F0) and (lowerModelPixel = 0x000000)) ; only lower on, turn on upper
	{
		quickClick(3shapeModelSlider["upperX"], 3shapeModelSlider["upperY"])
	}
	BlockInput MouseMoveOff
	return
}

Ortho_transparency()
{
	PixelGetColor, upperModelPixel, 3shapeModelSlider["upperX"], 3shapeModelSlider["upperY"]
	PixelGetColor, lowerModelPixel, 3shapeModelSlider["lowerX"], 3shapeModelSlider["lowerY"]
	PixelGetColor, upperTransPixel, 3shapeModelSlider["upperTransX"], 3shapeModelSlider["upperY"]
	PixelGetColor, lowerTransPixel, 3shapeModelSlider["lowerTransX"], 3shapeModelSlider["lowerY"]

	if (upperModelPixel = 0x000000) or (upperModelPixel = 0xF0F0F0) ; if upper model exists
	{
		if (upperTransPixel = 0xC9C3BB)
		{
			quickClick(3shapeModelSlider["upperTransX"], 3shapeModelSlider["upperY"])
		}
		Else
		{
			quickClick(3shapeModelSlider["upperHalfX"], 3shapeModelSlider["upperY"])
		}
	}
	return
}

Ortho_Wax(firstKnife, secondKnife, lastTick) ; Tool for swapping out wax knifes
{
	BlockInput MouseMove
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	ControlGetText, PrepTool, TdfGroupInfo1, ahk_group ThreeShape

	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular
	{
		BlockInput MouseMoveOff
		return
	}

	if PrepTool != "Wax knife settings"
		WinActivate ahk_group ThreeShape
		quickClick(3shapeButtons["waxKnifeX"], 3shapeButtons["waxKnifeY"])
	currentTick := A_TickCount
	if currentTick - lastTick > 900
	{
		Send %firstKnife%
	}
	else
	{
		Send %secondKnife%
	}
	BlockInput MouseMoveOff
	return currentTick
}

Ortho_takeBitePics(patientInfo) 
{
	BlockInput MouseMove

	if !WinExist("ahk_group ThreeShape")
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Wrong Window, Case must be open in OrthoAnalyzer or ApplianceDesigner
		Exit
	}

	WinActivate, ahk_group ThreeShape
	WinWaitActive, ahk_group ThreeShape

	FileRemoveDir, %screenshotDir%, 1
	FileCreateDir, %screenshotDir%
	screenshotName := patientInfo["firstName"] patientInfo["lastName"]

	quickClick(3shapeButtons["allViewX"], 3shapeButtons["frontViewY"])
	Sleep, 500
	CaptureScreen(screenshotName "Front.jpg") ; function defined in CaptureScreen Library

	quickClick(3shapeButtons["allViewX"], 3shapeButtons["leftViewY"])
	Sleep, 500

	CaptureScreen(screenshotName "Left.jpg")

	quickClick(3shapeButtons["allViewX"], 3shapeButtons["rightViewY"])
	Sleep, 500

	CaptureScreen(screenshotName "Right.jpg")

	BlockInput MouseMoveOff
	return screenshotDir
}

Ortho_toolSelect(selectedTool) ; select or apply specified tool
{
	ControlGetText, PrepStep, TdfInfoCaption2, ahk_group ThreeShape
	if PrepStep not in Sculpt Maxillary,Sculpt Mandibular 
	{
		return
	}
	ControlGetText, activeTool, TdfGroupInfo1, Feature

	WinActivate Feature, Sculpt toolkit ; the window inside 3Shape with the tool buttons on it

	if (selectedTool = "splineCut")
	{
		if (activeTool != "Spline cut settings")
		{
			quickClick(3shapeButtons["splineCutX"], 3shapeButtons["splineCutY"], 3shapeButtons["splineSmoothX"], 3shapeButtons["splineSmoothY"], 100)
		}
		Else
		{
			quickClick(3shapeButtons["splineCutApplyX"], 3shapeButtons["splineCutApplyY"])
		}
	}
	else if (selectedTool = "planeCut")
	{
		if (activeTool != "Plane cut settings")
		{
			quickClick(3shapeButtons["planeCutX"], 3shapeButtons["planeCutY"])
		}
		Else
		{
			quickClick(3shapeButtons["planeCutApplyX"], 3shapeButtons["planeCutApplyY"])
		}
	}
	else if (selectedTool = "artifact")
	{
		if (activeTool != "Remove Artifacts settings")
		{
			quickClick(3shapeButtons["artifactX"], 3shapeButtons["artifactY"])
		}
		Else
		{
			quickClick(3shapeButtons["artifactApplyX"], 3shapeButtons["artifactApplyY"])
		}
	}
	return 
}

Ortho_nextButton()
{
	WinActivate ahk_group ThreeShape
	ControlGetText, currentPrepStep, TdfInfoCaption2, ahk_group ThreeShape
	ControlGetText, checkBoxText, TCheckBox5, Virtual Base
	
	if (InStr("Virtual base", currentPrepStep) and (checkBoxText = "Decimate base to")) ; on the "fit base" step, set curve to 0
	{
		ortho_sendText("0", 3shapeFields["baseCurve"], "Virtual Base")
		WinActivate ahk_group ThreeShape
		quickClick(3shapeButtons["nextButtonX"], 3shapeButtons["nextButtonY"])
	}
	else if InStr("Prepare occlusion,Setup plane alignment,Sculpt maxillary,Sculpt mandibular,Virtual base", currentPrepStep)
	{
		quickClick(3shapeButtons["nextButtonX"], 3shapeButtons["nextButtonY"])
	}
	else if InStr("Finish", currentPrepStep)
	{
		quickClick(3shapeButtons["nextButtonX"], 3shapeButtons["nextButtonY"])
		return true
	}
	return false
}
