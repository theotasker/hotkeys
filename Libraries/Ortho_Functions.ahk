; ===========================================================================================================================
; Library for all Standalone functions pertaining to Ortho Analyzer and Appliance designer
; ===========================================================================================================================

; ===========================================================================================================================
; Constants and groups
; ===========================================================================================================================

; constants for 3Shape button positions
global allViewX := 1902
global topViewY := 286
global bottomViewY := 253
global rightViewY := 215
global leftViewY := 177
global frontViewY := 106
global backViewY := 140
global transparencyY := 800
global transparencyY := 800

global artifactX := 120
global artifacty := 173
global planeCutX := 165
global planeCutY := 175
global splineCutX := 205
global splineCutY := 175
global splineSmoothX := 195
global splineSmoothY := 300
global waxKnifeX := 35
global waxKnifeY := 175

global nextButtonX := 190
global nextButtonY := 30

; starting ticks for double-tap functions for 3D mouse
topTick := A_TickCount
sideTick := A_TickCount
frontTick := A_TickCount

waxOneTick := A_TickCount
waxTwoTick := A_TickCount
waxThreeTick := A_TickCount

; Groups to determine active windows
GroupAdd, ThreeShape, OrthoAnalyzer - [
GroupAdd, ThreeShape, ApplianceDesigner - [

GroupAdd, ThreeShapePatient, New patient model set info
GroupAdd, ThreeShapePatient, New patient info

GroupAdd, ThreeShapeExe, ahk_exe OrthoAnalyzer.exe
GroupAdd, ThreeShapeExe, ahk_exe ApplianceDesigner.exe



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

    WinActivate, ahk_exe OrthoAnalyzer.exe
    WinWaitActive, Open patient case,, 10

    if ErrorLevel
    {
		BlockInput MouseMoveOff
		Gui, Destroy
        MsgBox, Couldn't get focus on Ortho Analyzer
        return
    }
    Sleep, 200

    ControlFocus, Advanced search, Open patient case
    Sleep, 400
    Send, {Enter}
    Sleep, 400

	if (searchMethod = "patientName")
	{
		ortho_sendText(patientInfo["firstName"], "TEdit10", "Open patient case")

		ortho_sendText(patientInfo["lastName"], "TEdit9", "Open patient case")

		ortho_sendText(patientInfo["clinicName"], "Edit1", "Open patient case")
	}
	else 
	{
		ortho_sendText(patientInfo["scriptNumber"], "TEdit16", "Open patient case")
	}

	ControlFocus, TButton2, Open patient case
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
	quickClick("78", "46")

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
        BlockInput, MouseMove
		ortho_sendText(patientInfo["firstName"] patientInfo["lastName"], "TEdit4", "New patient info")
		ortho_sendText(patientInfo["firstName"], "TEdit10", "New patient info")
		ortho_sendText(patientInfo["lastName"], "TEdit9", "New patient info")
		ortho_sendText(patientInfo["clinicName"], "Edit1", "New patient info")
        BlockInput, MouseMoveOff
    }
    else if WinActive("New patient model set info")
    {
        ControlFocus, TEdit5
        Sleep, 100
        Send, % patientInfo["scriptNumber"]
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


Ortho_View(firstViewY, secondViewY, lastActionTick) ; clicks the view button declared before the function is called, swaps to a secondary view if pressed twice
{
    global toggle
	BlockInput MouseMove

    currentTick := A_TickCount
	if currentTick - lastActionTick > 1000
	{
		WinActivate ahk_group ThreeShape
		quickClick(allViewX, firstViewY)
		toggle := 1
	}
	else
	{
		if toggle = 1
		{
			WinActivate ahk_group ThreeShape
			quickClick(allViewX, secondViewY)

			toggle := 2
		}
		else if toggle = 2
		{
			WinActivate ahk_group ThreeShape
			quickClick(allViewX, firstViewY)
			toggle := 1
		}
	}
	BlockInput MouseMoveOff
	return currentTick
}

Ortho_VisibleModel() ; Swaps Visisble Model
{
	BlockInput MouseMove
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
	BlockInput MouseMoveOff
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
		quickClick(waxKnifeX, waxKnifeY)
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

Ortho_Export() ; While in model edit, clicks the green check mark and exports model
{
	BlockInput MouseMove
	if !WinExist("ahk_group ThreeShape")
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Wrong Window, Must be in a case to use this function
		return
	}

	BlockInput MouseMove
	WinActivate ahk_group ThreeShape
	Click, 1088, 95 ; Check Mark
	BlockInput MouseMoveOff

	WinWaitActive, Open patient case,, 30
	if ErrorLevel
	{
		BlockInput MouseMoveOff
        Gui, Destroy
		MsgBox,, Timeout Error, "Open Patient Case" window took too long to open
		Exit
	}
	
	BlockInput MouseMove ; Export Clicks
	Sleep, 200
	Click, 389, 46
	Sleep, 100
	Send {Down 6}
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

	BlockInput MouseMoveOff
	return
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

	FileRemoveDir, %screenshotDir%, 1
	FileCreateDir, %screenshotDir%

	screenshotName := patientInfo["firstName"] patientInfo["lastName"]

	Ortho_View(frontViewY, bottomViewY, topTick)
	Sleep, 500
	CaptureScreen(screenshotName "Front.jpg") ; function defined in CaptureScreen Library

	Ortho_View(leftViewY, bottomViewY, topTick)
	Sleep, 500

	CaptureScreen(screenshotName "Left.jpg")

	Ortho_View(rightViewY, bottomViewY, topTick)
	Sleep, 500

	CaptureScreen(screenshotName "Right.jpg")

	BlockInput MouseMoveOff
	return screenshotDir
}

ortho_sendText(to_send, target_box, target_window)
{
	While(return_text != to_send && loop_check < 5)
	{
		ControlFocus, %target_box%, %target_window%
		Sleep, 100
		Send, ^a
		Sleep, 100
		Send, {del}
		Sleep, 100
		Send, %to_send%
		Sleep, 100
		Send, {del}
		Sleep, 100
		ControlGetText, return_text, %target_box%, %target_window%
		loop_check++
	}
	If(return_text != to_send)
	{
		return False
	}
	else
	{
		return True
	}
}
