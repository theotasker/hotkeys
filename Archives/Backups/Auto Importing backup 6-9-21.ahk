; ===========================================================================================================================
; Fully Automated 3Shape Importing Workflow
; ===========================================================================================================================

; ===========================================================================================================================
; AHK stuff
; ===========================================================================================================================

#SingleInstance, Force

CoordMode, Mouse, Client
CoordMode, Pixel, Client

GroupAdd, ThreeShapePatient, New patient model set info ; Group used for waiting for popup
GroupAdd, ThreeShapePatient, New patient info

GroupAdd, AfterModelLoad, [\\APP03\  ; Group used after loading the model, checks if error message appears
GroupAdd, AfterModelLoad, Error

; ===========================================================================================================================
; Working Directories
; ===========================================================================================================================

InputDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Input\"
WorkDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Working\"
ExtraFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Extra Files\"
OrigFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Original Files\"
ErrorDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Errors\"
LogDir := "\\APP03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Logs\" ;folder for logs
LogFile := "\\APP03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" 

f6::
{
	Ortho_Restart()
	
	Loop
	{
		if FileExist(WorkDir "*")  ; Make sure the working folder is empty
		{
			ClearFolder(WorkDir, ExtraFileDir, "")
		}
		
		if FileExist(InputDir "*.stl")  ; Main Loop
		{
			Import_Status := Import()
			if (Import_Status != "Pass")
			{
				MsgBox % Import_Status
				ClearFolder(WorkDir, ErrorDir, Import_Status)
				Ortho_Restart()
			}
		}
		
		else
		{
			Sleep, 10000
		}
	}
}


Import()
{
	Global InputDir, WorkDir, ExtraFileDir, OrigFileDir
	
	; Importing files ======================================================================================================
	
	TempVar := 9999999999999999999  ; just larger than any date/time will be
	Loop files, %InputDir%*.stl  
	{
		if (A_LoopFileTimeCreated < TempVar)  ; goes through the directory, and replaces OldestFile if the current file is older
		{
			OldestFile := A_LoopFileName
			TempVar := A_LoopFileTimeModified
		}
	}
	
	Name_First := StrSplit(OldestFile, "~")[1]
	Name_Last := StrSplit(OldestFile, "~")[2]
	Name_Clinic := StrReplace(StrSplit(OldestFile, "~")[3], "_", " ")
	Number_Script := StrSplit(OldestFile, "~")[4]
	
	; For the oldest file found, check and see if there's an opposing arch
	if InStr(OldestFile, "~Upr[1]") ; only an upper expected
	{
		Filename_Upper := OldestFile
		Filename_Lower := "None"
	}
	else if InStr(OldestFile, "~Lwr[1]") ; only a lower expected
	{
		Filename_Upper := "None"
		Filename_Lower := OldestFile
	}
	else if InStr(OldestFile, "~Upr[2]") ; both expected
	{
		Filename_Upper := OldestFile
		Filename_Lower := StrReplace(OldestFile, "~Upr[2]", "~Lwr[2]")
	}
	else if InStr(OldestFile, "~Lwr[2]") ; both expected
	{
		Filename_Upper := StrReplace(OldestFile, "~Lwr[2]", "~Upr[2]")
		Filename_Lower := OldestFile
	}
	else
	{
		return "Bad file labels for " OldestFile
	}
	
	if (Filename_Upper != "None") 
	{
		if !FileExist(InputDir Filename_Upper) ; expected an upper but couldn't find it
		{
			return "Couldn't find upper file"
		}
		FileMove, %InputDir%%Filename_Upper%, %WorkDir%%Filename_Upper%
	}
	
	if (Filename_Lower != "None") 
	{
		if !FileExist(InputDir Filename_Lower) ; expected a lower but couldn't find it
		{
			return "Couldn't find lower file"
		}
		FileMove, %InputDir%%Filename_Lower%, %WorkDir%%Filename_Lower%
	}
	
	; Opening 3Shape ======================================================================================================
    if !WinExist("Open patient case")
    {
		Ortho_Restart()
		if !WinExist("Open patient case") ; if restarting didn't land on the open patient case window
		{
			return "Couldn't find open patient case window at the start"
		}
    }

    WinActivate, ahk_exe OrthoAnalyzer.exe
    WinWaitActive, Open patient case,, 10
    if ErrorLevel
    {
        return "Couldn't focus on Ortho analyzer at the start"
    }
    Sleep, 200
	
	; Search for existing patient ==========================================================================================
	
    ControlFocus, Advanced search, Open patient case
    Sleep, 200
    Send, {Enter}
    Sleep, 600
	
	if(Send_Text(Name_First, "TEdit8", "Open patient case") = "Fail")
	{
		return "Couldn't enter first name for patient search"
	}
	
	if(Send_Text(Name_Last, "TEdit7", "Open patient case") = "Fail")
	{
		return "Couldn't enter last name for patient search"
	}
	
	if(Send_Text(Name_Clinic, "Edit1", "Open patient case") = "Fail")
	{
		return "Couldn't enter clinic name for patient search"
	}
        
    ControlFocus, TButton2, Open patient case    
    Send {enter}
    Sleep, 100
    Send {down}

    if !WinExist("Open patient case")
    {
        return "Couldn't find open patient case window after entering search information"
    }
	
	Sleep, 3000
	
	; New patient/new model set =======================================================================================
    
        ; click the new patient button
    WinActivate Open patient case

    Click, 78, 46 ; clicks the new patient/model set button
    
    ; see which window opens
    WinWait, ahk_group ThreeShapePatient,, 10
	if ErrorLevel
    {
		WinActivate, ahk_group ThreeShapePatient
		WinWaitActive, ahk_group ThreeShapePatient,, 10
		if ErrorLevel
		{
			Click, 78, 46 ; clicks the noew patient/model set button
			WinWait, ahk_group ThreeShapePatient,, 10
			if ErrorLevel
			{
				return "Couldn't get focus patient info popup first time"
			}
			
		}
    }
    
	Sleep, 1000
	
    SetTitleMatchMode, 3
    if WinExist("New patient info",, model)  
    {
		if(Send_Text(Name_First " " Name_Last, "TEdit2", "New patient info") = "Fail")
		{
			return "Couldn't enter full name for patient creation"
		}

		if(Send_Text(Name_First, "TEdit8", "New patient info") = "Fail")
		{
			return "Couldn't enter first name for patient creation"
		}
		
		if(Send_Text(Name_Last, "TEdit7", "New patient info") = "Fail")
		{
			return "Couldn't enter last name for patient creation"
		}
		
		if(Send_Text(Name_Clinic, "Edit1", "New patient info") = "Fail")
		{
			return "Couldn't enter clinic name for patient creation"
		}
		
		ControlFocus, TButton2, New patient info
		Sleep, 200
		WinActivate, New patient info
		Send {return}
		Sleep, 2000
		
		SetTitleMatchMode, 3
		if WinExist("Warning")   ; Patient ID already exists, cancel out
		{
			ControlFocus, TButton1, Warning
			Sleep, 100
			send {return}
			Sleep, 100
			
			return "Patient already exists"
		}
			

		; click the new patient button
		WinActivate Open patient case
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, 78, 46
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
		
		; see which window opens
		WinWait, ahk_group ThreeShapePatient,, 5
		if ErrorLevel
		{
			return "Couldn't get focus patient info popup after entering patient info"
		}
	}
	
	Sleep, 1000
	
	if WinExist("New patient info",, model) ; second pass because why not
    {
		; function to paste patient info into new patient info fields using globals
		
		ControlFocus, TEdit2, New patient info
		Sleep, 200
		Send, %Name_First% %Name_Last%
		Sleep, 100
		
		ControlFocus, TEdit8, New patient info
		Sleep, 100
		Send, %Name_First%
		Sleep, 100
		
		ControlFocus, TEdit7, New patient info
		Sleep, 100
		Send, %Name_Last%
		Sleep, 100

		ControlFocus, Edit1, New patient info
		Sleep, 200
		Send, %Name_Clinic%
		
		ControlFocus, TButton2, New patient info
		Sleep, 200
		WinActivate, New patient info
		Send {return}
		Sleep, 2000
		
		SetTitleMatchMode, 3
		if WinExist("Warning")   ; Patient ID already exists, cancel out
		{
			ControlFocus, TButton1, Warning
			Sleep, 100
			send {return}
			Sleep, 100
			
			return "Patient already exists"
		}
			

		; click the new patient button
		WinActivate Open patient case
		BlockInput MouseMove
		MouseGetPos, x, y
		Click, 78, 46
		MouseMove, %x%, %y%, 0
		BlockInput MouseMoveOff
		
		; see which window opens
		WinWait, ahk_group ThreeShapePatient,, 5
		if ErrorLevel
		{
			return "Couldn't get focus patient info popup after entering patient info"
		}
	}
	
	SetTitleMatchMode, 3
	if !WinActive("New patient model set info") 
	{
		WinActivate, New patient model set info
		WinWaitActive, New patient model set info,, 5
		if ErrorLevel
		{
			return "Expected new model set window, couldn't find it"
		}
	}
	
	if(Send_Text(Number_Script, "TEdit5", "New patient model set info") = "Fail")
	{
		return "Couldn't enter clinic name for patient creation"
	}
	
	if (Filename_Upper = "None")  ; boxes for enabling upper/lower arches in new model set
		Click, 60, 90
	if (Filename_Lower = "None")
		Click, 60, 170
	
	Sleep, 500
	
	Send, {return}
	
	Sleep, 500
	
	if WinActive("Error")
	{
		; -------------------------------------------------------------error handling needed here
		return "Model set already exists"
	}
	
	WinWait, Open models,, 30  ; popup window after loading 3D environment
	if ErrorLevel
	{
		return "Popup for loading STLs didn't appear in time"
	}
	Sleep, 500
	
	WinActivate, Open models
	WinWaitActive, Open models,, 5
	if ErrorLevel
	{
		return "Couldn't activate open models window"
	}
	
	
	; Loading Models ==================================================================================================
	
	if (Filename_Upper != "None")
	{
		ControlFocus, TButton3, Open models   ; browse for file button
		Sleep, 200
		Send, {return}
		WinWaitActive, Open, , 10  ; directory search box
		if ErrorLevel
		{
			WinActivate, Open models
			WinWaitActive, Open models,, 5
			if ErrorLevel
			{
				return "Couldn't activate open models window"
			}
			ControlFocus, TButton3, Open models   ; browse for file button
			Sleep, 200
			Send, {return}
			WinWaitActive, Open, , 10  ; directory search box
			if ErrorLevel
			{
				return "Popup for selecting upper STL didn't appear in time"
			}
		}
		
		Sleep, 200
		ControlFocus, Edit1, Open
		Sleep, 200
		
		Send, %Filename_Upper%
		
		Sleep, 500
		
		Send, {return}
		
		WinWaitActive, Open models,, 10    ; back to the main popup
		if ErrorLevel
		{
			return "Didn't return to the main STL popup in time"
		}
		
		Sleep, 2000
		
		PixelGetColor, Pixel_Upper, 395, 50   ; after model is loaded, the box changes color
		
		If (Pixel_Upper != "0xFFFFFF")
		{
			Sleep, 2000
			
			PixelGetColor, Pixel_Upper, 395, 50   ; after model is loaded, the box changes color
			
			If (Pixel_Upper != "0xFFFFFF")
			{
				return "pixel for loaded upper model was not correct"
			}
			
		}
		
		
	}
	
	if (Filename_Lower != "None")
	{
		ControlFocus, TButton2, Open models   ; browse for file button
		Sleep, 200
		Send, {return}
		WinWaitActive, Open, , 10  ; directory search box
		if ErrorLevel
		{
			WinActivate, Open models
			WinWaitActive, Open models,, 5
			if ErrorLevel
			{
				return "Couldn't activate open models window"
			}
			ControlFocus, TButton2, Open models   ; browse for file button
			Sleep, 200
			Send, {return}
			WinWaitActive, Open, , 10  ; directory search box
			if ErrorLevel
			{
				return "Popup for selecting lower STL didn't appear in time"
			}
		}
		
		Sleep, 200
		ControlFocus, Edit1, Open
		Sleep, 200
		
		Send, %Filename_Lower%
		
		Sleep, 500
		
		Send, {return}
		
		WinWaitActive, Open models,, 10    ; back to the main popup
		if ErrorLevel
		{
			return "Didn't return to the main STL popup in time"
		}
		
		Sleep, 2000
		
		PixelGetColor, Pixel_Upper, 395, 250   ; after model is loaded, the box changes color
		
		If (Pixel_Upper != "0xFFFFFF") 
		{
			Sleep, 2000
		
			PixelGetColor, Pixel_Upper, 395, 250   ; after model is loaded, the box changes color
			
			If (Pixel_Upper != "0xFFFFFF")
			{
				return "pixel for loaded lower model was not correct"
			}
		}
		
		
	}
	
	
	Sleep, 1000
	
	Send, {space}   ; confirm loaded models
	
	
	SetTitleMatchMode, 2
	WinWait, ahk_group AfterModelLoad,, 30  ; waits for either the model to finish loading, or the error popup
	if ErrorLevel
	{
		return "Couldn't get either error popup or main window after loading model"
	}
	
	Sleep, 200
	
	SetTitleMatchMode, 3
	If WinExist("Error")  ; if 3Shape gives an error for the STLs, import anyways
	{
		WinActivate, "Error"
		ControlFocus, TButton1, Error
		sleep, 200
		Send, {Enter}
		SetTitleMatchMode, 2
		WinWait, [\\APP03\
		Sleep, 200
		SetTitleMatchMode, 3
	}
	
	Click, 30, 60 ; patient browser button
	
	WinWait, Open patient case,, 30
	if ErrorLevel
		{
			WinActivate Open patient case
			WinWaitActive, Open patient case,, 5
			if ErrorLevel
			{
				Click, 30, 60 ; patient browser button
				WinActivate Open patient case
				WinWaitActive, Open patient case,, 20
				if Errorlevel
				{
					return "Couldn't get back to the open patient case window"
				}
			}
		}

	Sleep, 200
        
    BlockInput, MouseMoveOff
	
	; Cleaning up ==================================================================================================
	
	try
	{
		if (Filename_Upper != "None")
		{
			FileMove, %WorkDir%%Filename_Upper%, %OrigFileDir%%Filename_Upper%
		}
		if (Filename_Lower != "None")
		{
			FileMove, %WorkDir%%Filename_Lower%, %OrigFileDir%%Filename_Lower%
		}
	}
	catch e
	{
		return "Couldn't move original files after importing"
	}
		

    return "Pass"
}

f9::
{
	Ortho_Restart()
	return
}

Ortho_Restart()
{
	While WinExist("ahk_exe OrthoAnalyzer.exe")
	{
		WinKill, ahk_exe OrthoAnalyzer.exe
		Sleep, 2000
	}
	
	Sleep, 5000
	
	if WinExist("ahk_exe OrthoAnalyzer.exe")
	{
		msgbox couldn't close orthoanalyzer
	}
	
	Run, OrthoAnalyzer.exe, \\app03\3Shape Ortho System\OrthoAnalyzer\
	
	WinWait, OrthoAboutForm,, 30
	WinWait, Open patient case,, 30
	return
}


ClearFolder(Directory, Repository, ErrorText)
{
	DirList := [] ; Array for listing files to pull from

	Loop Files, %Directory%* ; Appends each file to the list
		DirList.Push(A_LoopFileName)
	
	for key, filename in DirList
	{
		CurrentFilename := Directory filename
		ExportFilename := Repository StrSplit(filename, ".")[1] " (" ErrorText ")." StrSplit(filename, ".")[2]
		FileMove, %CurrentFilename%, %ExportFilename%, 1
	}
	return
}

Send_Text(to_send, target_box, target_window)
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
		ControlGetText, return_text, %target_box%, %target_window%
		loop_check++
	}
	If(return_text != to_send)
	{
		return "Fail"
	}
	else
	{
		return "Pass"
	}
}
