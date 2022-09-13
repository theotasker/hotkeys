; =========================================================================================================================
; Progress bar utilities and step settings
; =========================================================================================================================

; default progress bar location
global ProgX := "x788"
global ProgY := "y150"

; default step for the program
Step := "Prepping"
StartStep := "Start Digital Prep"
EndStep := "End Digital Prep"

Pause::
{
	SetBox()
	return
}

Progress: ; Set the location of the progress bar
{
	ProgX := "x788"
	ProgY := "y150"
	Gui, Destroy
	Gui, Add, Text, x0 y0 w320 h15 , Move this box to where you want your progress bar to be
	Gui, Add, Button, x52 y15 w120 h20 gSetLocation, Set Location
	Gui, Show, %ProgX% %ProgY% w320 h30, Select Location
	return
}

SetLocation: ; Subroutine for the progress bar location set GUI
{
	Gui, Show
	WinGetPos, VarX, VarY,,, Select Location
	ProgX := "x" + VarX
	ProgY := "y" + VarY
	Gui, Destroy
	return
}

Importing:
{
	Step := "Importing"
	StartStep := "Start Digital Import"
	EndStep := "End Digital Import"
	Gui, Destroy
	return
}

Prepping:
{
	Step := "Prepping"
	StartStep := "Start Digital Prep"
	EndStep := "End Digital Prep"
	Gui, Destroy
	Exit
}

Engraving:
{
	Step := "Engraving"
	StartStep := "Start Engraving"
	EndStep := "End Engraving"
	Gui, Destroy
	Exit
}

Printing:
{
	Step := "Printing"
	StartStep := "Start 3D Print Model"
	EndStep := "End 3D Print Model"
	Gui, Destroy
	Exit
}

GuiClose:
{
	Gui, Destroy
	Exit
}

ManagerCheckSub: ; runs when the check box is clicked, updates txt file and updates program variable for website
{
	Gui, Submit, NoHide

	if ManagerCheckVar = 1
	{
		FileDelete, %ManagerCheckFile%
		FileAppend, yes, %ManagerCheckFile%
		managercheck := "yes"
	}
	if ManagerCheckVar = 0
	{
		FileDelete, %ManagerCheckFile%
		FileAppend, no, %ManagerCheckFile%
		managercheck := "no"
	}
	return
}

Guide:
{
	Run, https://docs.google.com/document/d/1nQaEP69PKuJdn2lcOLDvR4fMiKJfQNFgxwkA2em04HM/edit?usp=sharing
	return
}

SetBox()
{
	global
	Gui, Add, Button, x12 y210 w100 h30 gImporting, Importing
	Gui, Add, Button, x112 y210 w100 h30 gPrepping, Prepping
	Gui, Add, Button, x212 y210 w100 h30 gEngraving, Engraving
	Gui, Add, Button, x312 y210 w100 h30 gPrinting, Printing
	Gui, Add, Button, x82 y10 w250 h30 gProgress, Change Progress Bar Location
	Gui, Add, Text, x82 y180 w250 h20 +Center, Change Current Step
	Gui, Add, Text, x5 y80 w250 h100, Cases imported = %importedlocal%`nCases prepped = %preppedlocal%`nCases engraved = %engravedlocal%`nHotkeys used = %hotkeyslocal%`nMouse actions saved = %clickslocal%`nKeystrokes saved = %strokeslocal%


	Gui, Font, underline
	Gui, Add, Text, cBlue gGuide x162 y120 w120 h14 +Center, Click here to view guide
	Gui, Show, w439 h253, NeoLab AutoHotKey Settings
	return
}

progressBar(action, percent)
{
	global
	if (action = "create")
	{
		Gui, Add, Progress, vprogress w300 h45
		Gui, Show, w320 h25 %ProgX% %ProgY%, Script Running
		Gui, +AlwaysOnTop
	}
	else if (action = "update")
	{
		GuiControl,, Progress, %percent%
	}
	else
	{
		GuiControl,, Progress, 100
		Gui, Destroy
	}
	return
}