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
InputDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Server\Input\"
WorkDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Server\Working\"
ExtraFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Server\Extra Files\"
OrigFileDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Server\Original Files\"
ErrorDir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Server\Errors\"
LogDir := "\\APP03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Server\Logs\" ;folder for logs
LogFile := "\\APP03\Scans\~Digital Dept Share\R&D Network\Auto Engraving\Server\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" 


; ===========================================================================================================================
; Main Loop
; ===========================================================================================================================

Netfabb_Restart()

Loop
{
	if FileExist(WorkDir "*")  ; Make sure the working folder is empty
	{
		ClearFolder(WorkDir, ExtraFileDir)
	}
	
	if FileExist(InputDir "*.stl")  ; Main Loop
	{
		filename := ImportFile()
		Sleep, 600
		
		NetfabbStatus := Netfabb_Engrave(filename)
	
		If(NetfabbStatus != "Pass")  ; If engraving errored out
		{
			FileRead, ExtantLog, %Logfile%
			If !InStr(ExtantLog, filename) ;  If the file isn't already in the log file
			{
				InputFilename := InputDir filename
				WorkFilename := WorkDir filename
				FileMove, %WorkFilename%, %InputFilename%  ; move it back to the input folder
			}
			ClearFolder(WorkDir, ErrorDir)    ; move everything else into the error folder
			Log_Update(filename, "Fail", NetfabbStatus)
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
				FileRead, ExtantLog, %Logfile%
				If !InStr(ExtantLog, filename) ;  If the file isn't already in the log file
				{
					InputFilename := InputDir filename
					WorkFilename := WorkDir filename
					FileMove, %WorkFilename%, %InputFilename%  ; move it back to the input folder
				}
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
	WinWaitActive, Autodesk Netfabb,, 10
	if ErrorLevel
	{
		return "Couldn't get focus on netfabb at start of engraving"
	}
	
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
	If(TopEdgeLeveling = "Error")
	{
		return "Error with pixel search for leveling, top edge"
	}
	
	BottomEdgeLeveling := Pixel_Search("y", 860, TopEdgeLeveling + 5, 20, "BG")
	If(BottomEdgeLeveling = "Error")
	{
		WinActivate, Autodesk Netfabb
		WinWaitActive, Autodesk Netfabb,, 10
		if ErrorLevel
		{
			return "Couldn't get focus on netfabb before retrying pixel search bottom edge level"
		}
		BottomEdgeLeveling := Pixel_Search("y", 860, TopEdgeLeveling + 5, 20, "BG")
		If(BottomEdgeLeveling = "Error")
		{
			return "Error with pixel search for leveling, bottom edge"
		}
	}
	
	If(BottomEdgeLeveling - TopEdgeLeveling < 100) ; repeats search if the height isn't big enough
	{
		BottomEdgeLeveling := Pixel_Search("y", 860, BottomEdgeLeveling + 5, 25, "BG")
		If(BottomEdgeLeveling = "Error")
		{
			BottomEdgeLeveling := Pixel_Search("y", 860, TopEdgeLeveling + 5, 20, "BG")
			If(BottomEdgeLeveling = "Error")
			{
				return "Error with pixel search for leveling, bottom edge after height check"
			}
		}
	}
	ModelCenterY := Round((TopEdgeLeveling + BottomEdgeLeveling)/2)
	
	BGPixel := 0x000000
	PixelGetColor, PixelColor, 860, ModelCenterY
	If(PixelColor = BGPixel)
	{
		return "Pixel for leveling was black, not on the model"
	}
	
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
	If(TopEdge = "Error")
	{
		WinActivate, Autodesk Netfabb
		WinWaitActive, Autodesk Netfabb,, 10
		if ErrorLevel
		{
			return "Couldn't get focus on netfabb before retrying pixel search for labeling"
		}
		TopEdge := Pixel_Search("y", 860, 75, 10, "MD")
		If(TopEdge = "Error")
		{
			return "Error with finding top edge during label placement"
		}
	}
	
	BottomEdge := Pixel_Search("y", 860, TopEdge + 5, 10, "BG")
	If(BottomEdge = "Error")
	{
		WinActivate, Autodesk Netfabb
		WinWaitActive, Autodesk Netfabb,, 10
		if ErrorLevel
		{
			return "Couldn't get focus on netfabb before retrying pixel search for labeling"
		}
		BottomEdge := Pixel_Search("y", 860, TopEdge + 5, 10, "BG")
		If(BottomEdge = "Error")
		{
			return "Error with finding bottom edge during label placement"
		}
	}
	
	If(BottomEdge - TopEdge < 200)  ; if the size was too small, continue search
	{
		BottomEdge := Pixel_Search("y", 860, BottomEdge + 10, 10, "BG")
	}
	
	If(TopEdge = "Error") or (BottomEdge = "Error")
	{
		return "Error with finding bottom edge while retrying to find bottom during labeling"
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
	If(LeftEdge = "Error") or (RightEdge = "Error")
	{
		WinActivate, Autodesk Netfabb
		WinWaitActive, Autodesk Netfabb,, 10
		if ErrorLevel
		{
			return "Couldn't get focus on netfabb before retrying pixel search for labeling"
		}
		LeftEdge := Pixel_Search("x", 860, EngravingYPosition, -10, "BG")
		RightEdge := Pixel_Search("x", 860, EngravingYPosition, 10, "BG")
		If(LeftEdge = "Error") or (RightEdge = "Error")
		{
			return "Error with finding left and right edges during label placement"
		}
	}
	
	
	EngravingXPosition := Round((LeftEdge + RightEdge)/2)
	
	BGPixel := 0x000000
	PixelGetColor, PixelColor, EngravingXPosition, EngravingYPosition
	If(PixelColor = BGPixel)
	{
		return "Pixel for engraving was black, not on the model"
	}
	
	; Click the calculated point to apply the labeling
	BlockInput MouseMove
    MouseGetPos, x, y
    Click, %EngravingXPosition%, %EngravingYPosition%
    MouseMove, %x%, %y%, 0
    BlockInput MouseMoveOff
	
	; added wait to attempt to fix error with missing the control focus send
	WinActivate, Autodesk Netfabb
	WinWaitActive, Autodesk Netfabb,, 10
	if ErrorLevel
	{
		return "Couldn't get focus on netfabb before applying the engraving"
	}
	
	; Hit the "Apply" button and confirm the "delete original model" popup
	ControlFocus, Apply, Autodesk Netfabb
	Sleep, 300 
	Send {Enter}
	WinWaitActive, Confirmation,, 4
	if ErrorLevel
	{
		return "Couldn't get focus on the 'delete original model' popup"
	}
	Sleep, 200
	Send {Enter}
	Sleep, 2000 ; waits for the engraving to finish
	
	; =========================================================================================================
	; Repair
	
	Send {Control Down}a{Control Up}  ; reselect the model
	Sleep, 100
	
	Send {alt} ; open automatic repair dialogue
	Sleep, 100
	Send p
	Sleep, 100
	Send e
	Sleep, 200
	
	WinWaitActive, Automatic Repair,, 10
	if ErrorLevel
	{
		WinActivate, Autodesk Netfabb
		WinWaitActive, Autodesk Netfabb,, 10
		if ErrorLevel
		{
			return "Couldn't get focus on netfabb before opening repair dialogue"
		}
		
		Send {esc} ; attempts to close any open windows
		Sleep, 200
		
		Send {Control Down}a{Control Up}  ; reselect the model
		Sleep, 100
		
		Send {alt} ; open automatic repair dialogue
		Sleep, 100
		Send p
		Sleep, 100
		Send e
		Sleep, 200
		
		WinWaitActive, Automatic Repair,, 10
		if ErrorLevel
		{
			return "Couldn't open automatic repair dialogue"
		}
	}
	
	Sleep, 200
	Send {Down 2}`t{Enter}
	WinWait, Autodesk Netfabb Standard, One job in queue, 10 ; waits for the repair job to enter que
	if ErrorLevel
	{
		return "Repair job didn't enter the queue"
	}
	Sleep, 1000
	
	WinWait, Autodesk Netfabb Standard, No jobs, 20 ; and waits for it to finish
	if ErrorLevel
	{
		return "Repair job didn't leave the queue"
	}
	
	Sleep, 100
	

	; =========================================================================================================
	; Export
	
	Send {Control Down}a{Control Up}  ; reselect the model
	Sleep, 200
	
	Send !f  ; open export 
	Sleep, 100
	Send r
	Sleep, 100
	
	Send {Down 2}{Enter}  ; gets to STL option
	WinWaitActive, Export,, 5
	if ErrorLevel
	{
		Send {esc} ; attempts to close any open windows
		Sleep, 200
		
		WinActivate, Autodesk Netfabb
		WinWaitActive, Autodesk Netfabb,, 10
		if ErrorLevel
		{
			return "Couldn't get focus on netfabb before opening export dialogue"
		}
		
		Send {Control Down}a{Control Up}  ; reselect the model
		Sleep, 200
		
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
	}

	Sleep, 200
	Send {Enter}
	Sleep, 200
	Send {Enter}
	Sleep, 200
	
	; Check to see if the file repair warning pops up
	Sleep, 400
	SetTitleMatchMode, 3
	if winexist("Export ")
	{
		return "The file was invalid at export"
	}
	SetTitleMatchMode, 1
	
	; Delete any remaining models
	WinActivate, Autodesk Netfabb 
	WinWaitActive, Autodesk Netfabb,, 10
	if ErrorLevel
	{
		return "Couldn't get focus netfabb before deleting models"
	}
	Send {Control Down}a{Control Up}
	Sleep, 100
	Send {Delete}
	Sleep, 100
	Send {Enter}
	Sleep, 800

	return "Pass"
}

; closes and reopens netfabb
Netfabb_Restart()
{
	LoopCount := 0
	Send {esc}
	Sleep, 200
	
	While WinExist("ahk_exe netfabb.exe")  ; close down netfabb
	{
		if WinExist("Warning")
		{
			WinActivate, Autodesk Netfabb
			Sleep, 500
			ControlFocus, Button2, Warning
			Sleep, 500
			Send {Enter}
			Sleep, 500
		}
		WinKill, ahk_exe netfabb.exe
		Sleep, 1000
		LoopCount+=1
		
		If LoopCount = 20
		{
			msgbox looped too many times
			Exit
		}
	}
	
	Sleep, 3000

	Run, netfabb.exe, C:\Program Files\Autodesk\Netfabb Standard 2019\
	WinWaitActive, ClicSplashScreenForm, , 10
	if ErrorLevel
	{
		Netfabb_Restart()
		return
	}
	WinWaitNotActive, ClicSplashScreenForm, , 10
	if ErrorLevel
	{
		Netfabb_Restart()
		return
	}
	
	Sleep, 2000

	WinWaitActive, ahk_exe netfabb.exe, , 20
	if ErrorLevel
	{
		If WinExist("ahk_exe netfabb.exe")
		{
			WinActivate, ahk_exe netfabb.exe
			WinWaitActive, ahk_exe netfabb.exe, , 10
			if ErrorLevel
			{
				Netfabb_Restart()
				return
			}
		}
		else
		{
			Netfabb_Restart()
			return
		}
	}
	Sleep, 2000
	
	SetTitleMatchMode, 3
	if WinExist("Restore backup projects")
	{
		WinClose, Restore backup projects
		Sleep, 500
	}
	SettitleMatchMode, 1
	return
}

; Searches axis from start point until pixel changes color and returns location		
Pixel_Search(XorY, StartingX, StartingY, Increment, BGorMD) 
{
	; Set background color and get starting pixel color
	BGPixel := 0x000000
	PixelGetColor, PixelColor, StartingX, StartingY
	
	BackupLoopBreak := 0

	If(XorY = "x") ; Searches x axis
	{
		Counter := StartingX
		
		If(BGorMD = "MD") ; searching for start of model color
		{
			While(PixelColor = BGPixel)
			{
				Counter := Counter + Increment
				PixelGetColor, PixelColor, Counter, StartingY
				If(Counter > 1400)
				{
					return "Error"
				}
				
				BackupLoopBreak+=1 ; trying to catch loop that doesn't quit on it's own
				If(BackupLoopBreak > 100)
				{
					Log_Update("loopbreak", "Fail", "pixelsearch faulted to backup loop break")
					return "Error"
				}
			}
			return Counter
		}
		Else If(BGorMD = "BG") ; Searching for start of background color
		{
			While(PixelColor != BGPixel)
			{
				Counter := Counter + Increment
				PixelGetColor, PixelColor, Counter, StartingY
				If(Counter > 1400)
				{
					return "Error"
				}
				
				BackupLoopBreak+=1 ; trying to catch loop that doesn't quit on it's own
				If(BackupLoopBreak > 100)
				{
					Log_Update("loopbreak", "Fail", "pixelsearch faulted to backup loop break")
					return "Error"
				}
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
				If(Counter > 800)
				{
					return "Error"
				}
				
				BackupLoopBreak+=1 ; trying to catch loop that doesn't quit on it's own
				If(BackupLoopBreak > 100)
				{
					Log_Update("loopbreak", "Fail", "pixelsearch faulted to backup loop break")
					return "Error"
				}
			}
			return Counter
		}
		Else If(BGorMD = "BG") ; Searching for start of background color
		{
			While(PixelColor != BGPixel)
			{
				Counter := Counter + Increment
				PixelGetColor, PixelColor, StartingX, Counter

				If(Counter > 1100)
				{
					return "Error"
				}
				
				BackupLoopBreak+=1 ; trying to catch loop that doesn't quit on it's own
				If(BackupLoopBreak > 500)
				{
					Log_Update("loopbreak", "Fail", "pixelsearch faulted to backup loop break")
					return "Error"
				}
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

	Loop Files, %Directory%* ; Appends each file to the list
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


