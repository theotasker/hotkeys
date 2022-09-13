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

OutputDir := "\\app03\Scans\~Digital Dept Share\~builds temp\~Prepped Models & Ready for Print\"
; OutputDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Output\"

InputDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Input\"
WorkDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Working\"
ExtraFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Extra Files\"
OrigFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Original Files\"
ErrorDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Errors\"
LogDir := "\\APP03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Logs\" ;folder for logs
LogFile := "\\APP03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" 


; ===========================================================================================================================
; Main Loop
; ===========================================================================================================================

f6::
{
	Netfabb_Restart()
	
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
				; MsgBox, 0, %NetfabbStatus%
				Log_Update(filename, "Fail", NetfabbStatus)
				ClearFolder(WorkDir, ErrorDir)
				Netfabb_Restart()
			}
			else  ; If engraving succeeded
			{
				if FileExist(WorkDir (StrSplit(filename, ".")[1]) " (labeled).stl")  ; making sure file exported properly
				{
					ExportFile(filename, OrigFileDir)
					FinishedFile := (StrSplit(filename, ".")[1]) " (labeled).stl"
					ExportFile(FinishedFile, OutputDir)
					Log_Update(filename, "Pass", NetfabbStatus)
				}
				else
				{
					; MsgBox, 0, Error, File was not properly exported
					Log_Update(filename, "Fail", "File was not properly exported")
					ClearFolder(WorkDir, ErrorDir)
					Netfabb_Restart()
				}
				
			}
		}
			
		else  ; if the input is empty, wait 10 seconds
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
	; Update Arch variable, reject if it doesn't contain that information
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
		return "Not labeled Upr or Lwr"
	}
	
	; Activates Netfabb
	WinActivate, Autodesk Netfabb
	WinWaitActive, Autodesk Netfabb
	
	; =========================================================================================================
	; Importing
	
	Send {alt}  ; open add parts dialogue
	Sleep, 100
	Send f
	Sleep, 100
	Send p
	WinWaitActive, Add Parts, , 2
	if ErrorLevel
	{
		return "Couldn't open add parts dialogue"
	}
	
	Sleep, 200  ; highlight model and import
	Send {Shift Down}`t`t{Shift Up}
	Sleep, 200
	Send {Right}
	Sleep, 200
	Send {Enter}
	WinWaitActive, CalculationThreadForm, , 4  ; this is the progress bar for importing an stl
	WinWaitClose, CalculationThreadForm, , 10
	if ErrorLevel
	{
		return "timout on model importing"
	}
	Sleep, 1000
	
	; =========================================================================================================
	; Check Model Dimensions
	
	ControlGetText, ModelLength, Edit9, Autodesk Netfabb Standard
	ControlGetText, ModelWidth, Edit8, Autodesk Netfabb Standard
	ControlGetText, ModelHeight, Edit7, Autodesk Netfabb Standard
	ModelLength := StrReplace(ModelLength, " mm", "")
	ModelWidth := StrReplace(ModelWidth, " mm", "")
	ModelHeight := StrReplace(ModelHeight, " mm", "")
	
	if(ModelWidth > (ModelHeight * 1.2)) or (ModelWidth > (ModelLength * 1.2))
	{
		return "Model width was greater than height or length, wrong orientation"
	}
	
	; =========================================================================================================
	; Leveling
	
	Send {Control Down}w{Control Up} ; activate the "align parts" module
	Sleep, 200
	
	if WinExist(ahk_exe netfabb.exe "Warning")
	{
		return "error starting leveling step"
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
	TopEdgeLeveling := Pixel_Search("y", 860, 75, 25, "MD")   ; Custom function at bottom of script
	BottomEdgeLeveling := Pixel_Search("y", 860, TopEdgeLeveling + 5, 25, "BG")
	If(BottomEdgeLeveling - TopEdgeLeveling < 100) ; repeats search if the height isn't big enough
	{
		BottomEdgeLeveling := Pixel_Search("y", 860, BottomEdgeLeveling + 5, 25, "BG")
	}
	ModelCenterY := Round((TopEdgeLeveling + BottomEdgeLeveling)/2)
	
    BlockInput MouseMove ; Click middle of model
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
	
	; =========================================================================================================
	; Spinning
	
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
	

	; =========================================================================================================
	; Labeling
	
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
	If(BottomEdge - TopEdge < 200)  ; if the size was too small, continue search
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
	
	; =========================================================================================================
	; Repair
	
	Send !pe ; gets to the repair module
	Sleep, 200
	
	Send {Down 2}`t{Enter}
	WinWait, Autodesk Netfabb Standard, One job in queue ; waits for the repair job to enter que
	Sleep, 1000
	
	WinWait, Autodesk Netfabb Standard, No jobs ; and waits for it to finish
	Send `t`t`t{Down}

	; =========================================================================================================
	; Export
	
	Send !f  ; open export 
	Sleep, 100
	Send r
	Sleep, 100
	
	Send {Down 2}{Enter}  ; gets to STL option
	WinWaitActive, Export,, 5
	if ErrorLevel
	{
		return "Couldn't open export dialogue"
	}
	else
	{
		Send {Enter 2}
	}
	
	; Delete any remaining models
	WinActivate, Autodesk Netfabb 
	WinWaitActive, Autodesk Netfabb
	Send {Control Down}a{Control Up}
	Sleep, 100
	Send {Delete}
	Sleep, 100
	Send {Enter}
	Sleep, 800

	return "Pass"
}

f7::
{
	Netfabb_Restart()
	return
}

; closes and reopens netfabb
Netfabb_Restart()
{
	if WinExist(Restore backup projects)
	{
		WinClose, Restore backup projects
		Sleep, 200
	}
	
	if WinExist("ahk_exe netfabb.exe")
	{
		WinClose, ahk_exe netfabb.exe
		Sleep, 200
	}
	
	if WinExist("ahk_exe netfabb.exe" Warning)
	{
		ControlFocus, Button2, Warning
		Send {Enter}
	}
	
	Sleep, 10000
	
	if WinExist("ahk_exe netfabb.exe")
	{
		WinClose, ahk_exe netfabb.exe
		Sleep, 200
	}
	
	if WinExist("ahk_exe netfabb.exe" Warning)
	{
		ControlFocus, Button2, Warning
		Send {Enter}
	}
	
	WinWaitClose, ahk_exe netfabb.exe, , 20
	
	Sleep, 5000
	
	if WinExist("ahk_exe netfabb.exe")
	{
		MsgBox Couldn't close Netfabb
		Exit
	}
	
	
	if !WinExist("ahk_exe netfabb.exe")
	{
		Run, netfabb.exe, C:\Program Files\Autodesk\Netfabb Standard 2019\
		WinWaitActive, Autodesk Netfabb Standard, , 15
		if ErrorLevel
		{
			MsgBox, , Error, Couldn't reopen Netfabb
			Exit
			
			
		}
		Sleep, 2000
		
	}
	
	if WinExist(Restore backup projects)
	{
		WinClose, Restore backup projects
		Sleep, 200
	}
	
	return
}

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
		FileMove, %CurrentFilename%, %ExportFilename%, 1
	}
	return
}


ImportFile()
{
	global InputDir, WorkDir
	
	TempVar := 9999999999999999999  ; just larger than any date/time will be
	Loop files, %InputDir%*.stl  
	{
		if (A_LoopFileTimeCreated < TempVar)  ; goes through the directory, and replaces OldestFile if the current file is older
		{
			OldestFile := A_LoopFileName
			TempVar := A_LoopFileTimeModified
		}
	}
	
	InputFilename := InputDir OldestFile
	WorkFilename := WorkDir OldestFile

	FileMove, %InputFilename%, %WorkFilename%
	
	return OldestFile
}

ExportFile(filename, exportdir)
{
	global WorkDir
	
	CleanFileName := StrReplace(filename, " (labeled)", "")
	
	WorkFilename := WorkDir filename
	ExportFilename := exportdir CleanFileName
	FileMove, %WorkFilename%, %ExportFilename%, 1
	
	return
}

; ===========================================================================================================================
; Log Functions
; ===========================================================================================================================

Log_Update(FileName, ErrorStatus, ErrorText)
{
	global LogDir, LogFile
	; sets the text in case it has to create a new txt file
	LogText := "Files processed = 0`nFiles with errors = 0`n--------------------`n"

	if !FileExist(LogDir)
	{
	  FileCreateDir, %LogDir%
	}
	if !FileExist(LogFile)
	{
	  FileAppend, %LogText%, %Logfile% ; creates a new log file with the current variables
	}
	
	FileRead, ExtantLog, %Logfile%
	LogArray := StrSplit(ExtantLog, "`n") ; creates an array with an item for each line

	FilesProcessed := GetNumber(LogArray[1], "Files processed = ") ; converts to just the number
	FileErrors := GetNumber(LogArray[2], "Files with errors = ")
	
	if(ErrorStatus = "Fail")
	{
		FileErrors+=1
		ExtantLog := ExtantLog "`n" FileName " " ErrorText
	}
	else if(ErrorStatus = "Pass")
	{
		FilesProcessed+=1
	}
	
	AfterError := StrSplit(ExtantLog, "`n--------------------")
	UpdatedLog := "Files processed = " FilesProcessed "`nFiles with errors = " FileErrors "`n--------------------" AfterError[2]
	
	FileDelete, %Logfile%
	FileAppend, %UpdatedLog%, %Logfile%
	
	return
}

GetNumber(Tochange = "", ToRemove = "")
{
   Result := StrReplace(ToChange, ToRemove, "") ; removes the leading text
   Result := StrReplace(Result, "`n", "") ; removes the new line markers
   Result += 0 ; makes the number an integer
   return Result
}

Esc::Exit ; stops all processes