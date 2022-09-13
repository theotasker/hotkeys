; ===========================================================================================================================
; Library for all functions pertaining to Ortho Analyzer and Appliance Designer
; ===========================================================================================================================

; ===========================================================================================================================
; Import Functions
; ===========================================================================================================================

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
    
    ControlFocus, TEdit8, Open patient case    
    Sleep, 100
    Send, %firstname%
    Sleep, 100
    
    GuiControl,, Progress, 40
     
    ControlFocus, TEdit7, Open patient case    
    Sleep, 100
    Send, %lastname%
    Sleep, 100
    
    GuiControl,, Progress, 50

    ControlFocus, Edit1, Open patient case    
    Sleep, 100
    Send, %clinic%
    Sleep, 100
    
    GuiControl,, Progress, 60
        
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
    
    ControlFocus, TEdit2, New patient info
    Sleep, 200
    Send, %firstname% %lastname%
    Sleep, 100
    
    ControlFocus, TEdit8, New patient info
    Sleep, 100
    Send, %firstname%
    Sleep, 100
    
    ControlFocus, TEdit7, New patient info
    Sleep, 100
    Send, %lastname%
    Sleep, 100

    ControlFocus, Edit1, New patient info
    Sleep, 200
    Send, %clinic%

    BlockInput MouseMoveOff
    return
}

; ===========================================================================================================================
; Prepping Functions
; ===========================================================================================================================

; Creates group to select either ortho or appliance designer
GroupAdd, ThreeShape, OrthoAnalyzer - [
GroupAdd, ThreeShape, ApplianceDesigner - [

GroupAdd, ThreeShapePatient, New patient model set info
GroupAdd, ThreeShapePatient, New patient info

GroupAdd, ThreeShapeExe, ahk_exe OrthoAnalyzer.exe
GroupAdd, ThreeShapeExe, ahk_exe ApplianceDesigner.exe

; sets start variables for use in the double button tap functions
Var1 := A_TickCount
Var2 := A_TickCount
Var3 := A_TickCount
VarT := A_TickCount
VarR := A_TickCount
VarF := A_TickCount

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
	
	ControlFocus, TEdit13, Open patient case    
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
    
    SetTitleMatchMode, 3
	IfWinNotExist Sculpt, Wax knife settings
	{
		WinActivate Sculpt
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, 34, 105   ; click the wax knife box
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
	}
	startvar := A_TickCount
	if startvar - RefVar > 300
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