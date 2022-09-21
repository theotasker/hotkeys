; =========================================================================================================================
; Progress bar utilities and step settings
; =========================================================================================================================

Pause::
{
	SetBox()
	return
}

Progress: ; Set the location of the progress bar
{
	progressBarX := "x788"
	progressBarY := "y150"
	Gui, Destroy
	Gui, Add, Text, x0 y0 w320 h15 , Move this box to where you want your progress bar to be
	Gui, Add, Button, x52 y15 w120 h20 gSetLocation, Set Location
	Gui, Show, %progressBarX% %progressBarY% w320 h30, Select Location
	return
}

SetLocation: ; Subroutine for the progress bar location set GUI
{
	Gui, Show
	WinGetPos, VarX, VarY,,, Select Location
	progressBarX := "x" + VarX
	progressBarY := "y" + VarY
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

Gui_progressBar(action, percent:=0)
{
	global
	if (action = "create")
	{
		Gui, Add, Progress, vprogress w300 h45
		Gui, Show, w320 h25 %progressBarX% %progressBarY%, Script Running
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

Gui_finishImport(arches) 
{
	global finishOptions := {"upper":False, "lower":False,"arches":False, "auto":False}
	if (arches = "both")
	{
		Gui, Add, Text, x30 y20 w300 h14 +Center, Manual Import:
		Gui, Add, Button, x12 y40 w100 h30 gUpperManual, Upper Manual
		Gui, Add, Button, x112 y40 w100 h30 gLowerManual, Lower Manual
		Gui, Add, Button, x212 y40 w100 h30 gBothManual, Both Manual

		Gui, Add, Text, x30 y120 w300 h14 +Center, Auto Import:
		Gui, Add, Button, x12 y140 w100 h30 gUpperAuto, Upper Auto
		Gui, Add, Button, x112 y140 w100 h30 gLowerAuto, Lower Auto
		Gui, Add, Button, x212 y140 w100 h30 gBothAuto, Both Auto

		Gui, Show, w439 h253, Arch Selection (Two Arches Detected)

		WinWaitClose, Arch Selection (Two Arches Detected)
	}

	if (arches = "upper")
	{
		Gui, Add, Text, x30 y20 w300 h14 +Center, Manual Import:
		Gui, Add, Button, x12 y40 w100 h30 gUpperManual, Upper Manual

		Gui, Add, Text, x30 y120 w300 h14 +Center, Auto Import:
		Gui, Add, Button, x12 y140 w100 h30 gUpperAuto, Upper Auto

		Gui, Show, w439 h253, Arch Selection (Upper Arch Detected)

		WinWaitClose, Arch Selection (Upper Arch Detected)
	}

	if (arches = "lower")
	{
		Gui, Add, Text, x30 y20 w300 h14 +Center, Manual Import:
		Gui, Add, Button, x112 y40 w100 h30 gLowerManual, Lower Manual

		Gui, Add, Text, x30 y120 w300 h14 +Center, Auto Import:
		Gui, Add, Button, x112 y140 w100 h30 gLowerAuto, Lower Auto

		Gui, Show, w439 h253, Arch Selection (Lower Arch Detected)

		WinWaitClose, Arch Selection (Lower Arch Detected)
	}
	
	return finishOptions
}

UpperManual:
{
	finishOptions["upper"] := True
	finishOptions["auto"] := False
	Gui, Destroy
	return
}

LowerManual:
{
	finishOptions["lower"] := True
	finishOptions["auto"] := False
	Gui, Destroy
	return
}

BothManual:
{
	finishOptions["upper"] := True
	finishOptions["lower"] := True
	finishOptions["arches"] := True
	finishOptions["auto"] := False
	Gui, Destroy
	return
}

UpperAuto:
{
	finishOptions["upper"] := True
	finishOptions["auto"] := True
	Gui, Destroy
	return
}

LowerAuto:
{
	finishOptions["lower"] := True
	finishOptions["auto"] := True
	Gui, Destroy
	return
}

BothAuto:
{
	finishOptions["upper"] := True
	finishOptions["lower"] := True
	finishOptions["arches"] := True
	finishOptions["auto"] := True
	Gui, Destroy
	return
}