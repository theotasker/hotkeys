; ===========================================================================================================================
; Fully Automated Netfabb Engraving Workflow
; ===========================================================================================================================

; ===========================================================================================================================
; AHK stuff
; ===========================================================================================================================

#SingleInstance, Force

CoordMode, Mouse, Client
CoordMode, Pixel, Client

; ===========================================================================================================================
; Working Directories
; ===========================================================================================================================

InputDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Input\"
OutputDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Output\"
WorkDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Working\"
ExtraFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Extra Files\"
OrigFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Original Files\"
ErrorDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Errors\"


; ===========================================================================================================================
; Main Loop
; ===========================================================================================================================

f6::
{
	Loop
	{
		if FileExist(WorkDir "*.stl")  ; Make sure the working folder is empty
		{
			ClearFolder(WorkDir, ExtraFileDir)
		}
		
		if FileExist(InputDir "*.stl")  ; Main Loop
		{
			filename := ImportFile()
			
			Sleep, 400
			
			NetfabbStatus := Netfabb_Engrave(filename)
		
			If(NetfabbStatus != "Pass")  ; If engraving errored out
			{
				MsgBox, 0, Error, Engraving encountered an error
				ClearFolder(WorkDir, ErrorDir)
			}
			else  ; If engraving succeeded
			{
				if FileExist(WorkDir (StrSplit(filename, ".")[1]) " (labeled).stl")  ; making sure file exported properly
				{
					ExportFile(filename, OrigFileDir)
					FinishedFile := (StrSplit(filename, ".")[1]) " (labeled).stl"
					ExportFile(FinishedFile, OutputDir)
				}
				else
				{
					MsgBox, 0, Error, File was not properly exported
					ClearFolder(WorkDir, ErrorDir)
				}
				
			}
		}
			
		else
		{
			Sleep, 10000
		}
		
	}
}

; ===========================================================================================================================
; Netfabb Functions
; ===========================================================================================================================

; Main function
Netfabb_Engrave(filename) 
{	
	; Update arch, reject if it doesn't contain that information
	If InStr(filename, " Upr") or InStr(filename, " UBP") or InStr(filename, " UT")
	{
		Arch := "Upr"
	}
	Else If InStr(filename, " Lwr") or InStr(filename, " LBP") or InStr(filename, " LT")
	{
		Arch := "Lwr"
	}
	else
	{
		MsgBox, 0, Error, File must be labeled as "Upr" or "Lwr"
		return "Error"
	}
	
	; Activates Netfabb
	WinActivate, Autodesk Netfabb
	WinWaitActive, Autodesk Netfabb
	
	Netfabb_Import() ; opens "add model" dialogue and imports STL
	
	If(Netfabb_Level(Arch) = "Error") ; opens "align parts" module, clicks bottom, and applies
	{
		return "Error"
	}
	
	Sleep, 300
	Send {tab}{tab}{tab}{enter}
	
	; Spin the arch to ready for labelling
	If(Arch = "Upr")
	{
		Send {z 7}
	}
	Else If(Arch = "Lwr")
	{
		Send +{z 5}
	}
	Sleep, 200
	
	If(Netfabb_Label() = "Error") ; opens "text label" module, views bottom, calculates best location, clicks it, and applies
	{
		return "Error"
	}
	
	Netfabb_Repair() ; runs the "extended repair" on selected model
	
	Netfabb_Export() ; goes through the "export part" dialogue
	
	Netfabb_Clear() ; deletes any remaining models

	return "Pass"
}

Netfabb_Import() ; opens "add model" dialogue and imports STL
{
	Send {alt}
	Sleep, 100
	Send f
	Sleep, 100
	Send p
	WinWaitActive, Add Parts, , 2
	if ErrorLevel
	{
		Msgbox Couldn't open add parts
		Exit
	}
	Sleep, 200
	Send {Shift Down}`t`t{Shift Up}
	Sleep, 200
	Send {Right}
	Sleep, 200
	Send {Enter}
	WinWaitActive, CalculationThreadForm, , 4
	WinWaitClose, CalculationThreadForm, , 10
	if ErrorLevel
	{
		MsgBox Didn't import file properly
		Exit
	}
	Sleep, 1000
	return
}

Netfabb_Level(Arch)
{
	Send {Control Down}w{Control Up} ; activate the "align parts" module
	Sleep, 200
	
	if WinExist(ahk_exe netfabb.exe "Warning")
	{
		Exit
	}
	
	; Open view dialogue
	Send {Alt}
	Sleep, 100
	Send v
	Sleep, 100
	Send p
	Sleep, 100
	
	; Decide on front or back view
	If(Arch = "Upr")
	{
		Send b
		Sleep, 100
		Send {Enter}
	}
	Else if(Arch = "Lwr")
	{
		Send f
	}
	Sleep, 800
	
	; Search for top and bottom edges, find middle of model 
	TopEdgeLeveling := Pixel_Search("y", 860, 75, 25, "MD")
	BottomEdgeLeveling := Pixel_Search("y", 860, TopEdgeLeveling + 5, 25, "BG")
		; Check to make sure there was enough room to ensure a good click
	If(BottomEdgeLeveling - TopEdgeLeveling < 100)
	{
		BottomEdgeLeveling := Pixel_Search("y", 860, BottomEdgeLeveling + 5, 25, "BG")
	}
	
	
	ModelCenterY := Round((TopEdgeLeveling + BottomEdgeLeveling)/2)
	

	
	; Click middle of model
    BlockInput MouseMove
    MouseGetPos, x, y
    Click, 860, %ModelCenterY%
    MouseMove, %x%, %y%, 0
    BlockInput MouseMoveOff
	
	; Apply the leveling
	Sleep, 300
	ControlFocus, Apply, Autodesk Netfabb ; hits the apply button on the leveling step
	Sleep, 100
	Send {Enter}
	Sleep, 200
	ControlFocus, Window22, ahk_exe netfabb.exe
	Sleep, 100
	return
}

Netfabb_Label()
{
	Send {Control Down}e{Control Up} ; activate the "text label" module
	Sleep, 200
	
	; set view to bottom
	Send {Alt}
	Sleep, 100
	Send v
	Sleep, 100
	Send p
	Sleep, 100
	Send b
	Sleep, 100
	Send b
	Sleep, 100
	Send {Enter}
	Sleep, 400
	
	; Remove Man/Max from labelling text
	EditFields := ["Edit1", "Edit4", "Edit15", "Edit17", "Edit18"]
	For key, field in EditFields
	{
		ControlGetText, FieldContents, %field%, Autodesk Netfabb
		If InStr(FieldContents, "Upr") or InStr(FieldContents, "Lwr") or InStr(FieldContents, "UBP") or InStr(FieldContents, "LBP")
		{
			ControlFocus, %field%, Autodesk Netfabb
			Sleep, 100
			Send {End}
			Sleep, 100
			Send `b`b`b`b
		}
	}

	; Search for top and bottom edges
	TopEdge := Pixel_Search("y", 860, 75, 10, "MD")
	BottomEdge := Pixel_Search("y", 860, TopEdge + 5, 10, "BG")
	
	; Check to make sure there was enough room to ensure a good click, if not, keeps searching down
	If(BottomEdge - TopEdge < 200)
	{
		BottomEdge := Pixel_Search("y", 860, BottomEdge + 10, 10, "BG")
	}
	
	; decides on Y placement for full vs minimal base models
	If(BottomEdge < 500)
	{
		EngravingYPosition := Round(BottomEdge - ((BottomEdge - TopEdge)/3))
	}
	Else
	{
		EngravingYPosition := Round(TopEdge + ((BottomEdge - TopEdge)/6))
	}
	
	; From Y position, search for left and right edges and find center
	LeftEdge := Pixel_Search("x", 860, EngravingYPosition, -10, "BG")
	RightEdge := Pixel_Search("x", 860, EngravingYPosition, 10, "BG")
	EngravingXPosition := Round((LeftEdge + RightEdge)/2)
	
	; Click the calculated point to apply the labeling
	BlockInput MouseMove
    MouseGetPos, x, y
    Click, %EngravingXPosition%, %EngravingYPosition%
    MouseMove, %x%, %y%, 0
    BlockInput MouseMoveOff
	
	; added wait to attempt to fix error with missing the control focus send
	WinActivate, Autodesk Netfabb
	WinWaitActive, Autodesk Netfabb
	
	; Hit the "Apply" button and confirm the "delete original model" popup
	ControlFocus, Apply, Autodesk Netfabb
	Sleep, 100 
	Send {Enter}
	WinWaitActive, Confirmation,, 2 
	Send {Enter}
	Sleep, 2000 ; waits for the engraving to finish
	return
}

Netfabb_Repair()
{
	Send !pe ; gets to the repair module
	Sleep, 200
	
	Send {Down 2}`t{Enter}
	WinWait, Autodesk Netfabb Standard, One job in queue ; waits for the repair job to enter que
	Sleep, 1000
	
	WinWait, Autodesk Netfabb Standard, No jobs ; and waits for it to finish
	Send `t`t`t{Down}
	return
}

Netfabb_Export()
{
	Send !f
	Sleep, 100
	Send r
	Sleep, 100
	
	Send {Down 2}{Enter}
	WinWaitActive, Export,, 1
	if ErrorLevel
	{
		BlockInput MouseMoveOff
		MsgBox Timeout
		return
	}
	else
	{
		Send {Enter 2}
		BlockInput MouseMoveOff
	}
	return
}

Netfabb_Clear()
{
	WinActivate, Autodesk Netfabb
	WinWaitActive, Autodesk Netfabb
	Send {Control Down}a{Control Up}
	Sleep, 100
	Send {Delete}
	Sleep, 100
	Send {Enter}
	Sleep, 800
}

; ===========================================================================================================================
; Non-Specific Functions
; ===========================================================================================================================

; Searches axis from start point until pixel changes color and returns location		
Pixel_Search(XorY, StartingX, StartingY, Increment, BGorMD) 
{
	; Set background color and get starting pixel color
	BGPixel := 0x000000
	PixelGetColor, PixelColor, StartingX, StartingY

	If(XorY = "x") ; Searches x axis
	{
		Counter := StartingX
		
		If(BGorMD = "MD") ; searching for start of model color
		{
			While(PixelColor = BGPixel)
			{
				Counter := Counter + Increment
				PixelGetColor, PixelColor, Counter, StartingY
			}
			return Counter
		}
		Else If(BGorMD = "BG") ; Searching for start of background color
		{
			While(PixelColor != BGPixel)
			{
				Counter := Counter + Increment
				PixelGetColor, PixelColor, Counter, StartingY
			}
			return Counter
		}
	}
	Else If(XorY = "y") ; Searches y axis
	{
		Counter := StartingY
	
		If(BGorMD = "MD") ; searching for start of model color
		{
			While(PixelColor = BGPixel)
			{
				Counter := Counter + Increment
				PixelGetColor, PixelColor, StartingX, Counter
			}
			return Counter
		}
		Else If(BGorMD = "BG") ; Searching for start of background color
		{
			While(PixelColor != BGPixel)
			{
				Counter := Counter + Increment
				PixelGetColor, PixelColor, StartingX, Counter
			}
			return Counter
		}
	}
}

; ===========================================================================================================================
; File Handling Functions
; ===========================================================================================================================

ClearFolder(Directory, Repository)
{
	DirList := [] ; Array for listing files to pull from

	Loop Files, %Directory%*.stl ; Appends each file to the list
		DirList.Push(A_LoopFileName)
	
	for key, filename in DirList
	{
		CurrentFilename := Directory filename
		ExportFilename := Repository filename
		FileMove, %CurrentFilename%, %ExportFilename%
	}
	return
}


ImportFile()
{
	global InputDir, WorkDir
	
	InputDirList := [] ; Array for listing files to pull from

	Loop Files, %InputDir%*.stl ; Appends each file to the list
		InputDirList.Push(A_LoopFileName)
	
	InputFilename := InputDir InputDirList[1]
	WorkFilename := WorkDir InputDirList[1]

	FileMove, %InputFilename%, %WorkFilename%
	
	return InputDirList[1]
}

ExportFile(filename, exportdir)
{
	global WorkDir
	
	CleanFileName := StrReplace(filename, " (labeled)", "")
	
	WorkFilename := WorkDir filename
	ExportFilename := exportdir CleanFileName
	FileMove, %WorkFilename%, %ExportFilename%
	
	return
}