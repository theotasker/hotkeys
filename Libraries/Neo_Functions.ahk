neo_startWebDriver()
{
	global

	if WinExist("ahk_exe chrome.exe")
		MsgBox, 4, Close Chrome?, Close all current instances of Chrome? (suggested)
			IfMsgBox, Yes
				; Close all instances of chrome
				While WinExist("ahk_exe chrome.exe")
				{
					Loop, 10
					{
						WinClose, ahk_exe chrome.exe
					}
				}
			IfMsgBox, No
				nothing := 0

	; Check to make sure that the modified Chrome shortcut is in the right place
	chromelocation = %A_MyDocuments%\Automation\ChromeForAHK.lnk
	if !FileExist(chromelocation)
	{
		MsgBox,, Chrome Shortcut Not Found, Can't find the proper shortcut at %A_MyDocuments%\Automation\ChromeForAHK.lnk `nRemember to modify the shortcut to run in debug mode by adding "--remote-debugging-port=9222"
		Exit
	}

	; Open Chrome instance and assign it to neolab
	Run, ChromeForAHK.lnk, %A_MyDocuments%\Automation\
	Sleep, 500
	NeoDriver := ChromeGet()
	NeoDriver.Get("https://portal.rxwizard.com/cases")

	NeoCheck := true

	Exit
}

; function to check that the NeoDriver is still open
neo_stillOpen()
{
	global NeoDriver, NeoCheck

	if (NeoCheck != true)
		MsgBox, 4, Open NeoDriver?, No instance of NeoDriver found, initiate?
			IfMsgBox, Yes
				neo_startWebDriver()
			IfMsgBox, No
				Exit

	try temp := InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	catch e
	{
		MsgBox, 4, Webdriver Error, The tab for driving Portal.RXWizard was closed, initiate webdriver?
			IfMsgBox, Yes
				neo_startWebDriver()
			IfMsgBox, No
			{
				Gui, Destroy
				Exit
			}
		Gui, Destroy
		return
	}
}

; function to bring NEOLab website to the front
neo_activate()
{
	if WinExist("New England Orthodontic Laboratory - Google Chrome")
		WinActivate
	else
	{
		MsgBox,, Website Error, RxWizard Portal should be the top tab in its own instance
		exit
	}
	return
}

; function for returning to the cases page and entering the search field
neo_navigateToCases()
{
    global NeoDriver, Path_ScanScriptCSS

	Neo_StillOpen()

    if !InStr(NeoDriver.Url, "https://portal.rxwizard.com") ; Must be on a rxwizard page
    {
		Gui, Destroy
        MsgBox,, Wrong Page, Must be on an RxWizard Page
        Exit
    }

	Neo_Activate()

	Send {tab}
	Sleep, 100

    NeoDriver.findElementByCss(Path_ScanScriptCSS).click()

}

; function for getting to the review page from the edit page
neo_navigateReviewFromEdit()
{
    global NeoDriver, Path_ReviewButtonCSS

	Neo_StillOpen()

	Neo_Activate()

	; If on the edit page, click on the "review" button on the bottom, then either wait for the review page or
	; click the "confirm" button on the "over model count" popup.
    if InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	{
        try NeoDriver.findElementByCss(Path_ReviewButtonCSS).sendKeys(Keys.RETURN)
		catch e
		{
			MsgBox,, Couldn't Find Element, Couldn't find the "Review" button
			Gui, Destroy
			Exit
		}

	}
	else if InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")
		do = 0
	else
		MsgBox,, Wrong Page, Must be on the edit page for this function
	return

}

; function for hitting start/stop for %step% while already on the review page
neo_start()
{
	global NeoDriver, currentstep, currentstepxpath

	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/") ; checks to ensure on the review page
	{
		MsgBox,, Wrong Page, Must be on the review page for this function
		Gui, Destroy
		Exit
	}

	; Wait until the button for pushing step is clickable
	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.CSS_SELECTOR, currentstepxpath))
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the start button before it was clicked the first time
		Exit
	}

	; Try to find the text label on the button for pushing steps
	try webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't get the text from the start stop button before starting
		Exit
	}

	; check the retrieved button text against the currenstep variable
	if webstep = %currentstep%
	{
		try NeoDriver.findElementByCss(currentstepxpath).Click
		catch e
		{
			Gui, Destroy
			Msgbox,, Web Error, Couldn't click on the button at the start step
			Exit
		}

		Neo_Activate()

		WinWaitActive, New England Orthodontic Laboratory - Google Chrome,, 10

		Sleep, 200

		Send, {Enter}
	}
	else
	{
		Gui, Destroy
		MsgBox,, Web Error, Button wasn't on the right step
		Exit
	}
	return
}

; new function for stopping step, will replace the original
neo_Stop()
{
	global NeoDriver, currentstep, currentstepxpath

	; Wait until the button for pushing step is clickable
	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.CSS_SELECTOR, currentstepxpath))
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the stop button before it was clicked
		Exit
	}

	; Try to find the text label on the button for pushing steps
	try webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't get the text from the stop button before stopping
		Exit
	}

	; loop to check if the button is the right step, clicks it if it is, returns for 5 200ms loops if not
	failcount := 0
	Loop
	{
		if webstep = %currentstep%
		{
			NeoDriver.findElementByCss(currentstepxpath).Click

			Neo_Activate()

			WinWaitActive, New England Orthodontic Laboratory - Google Chrome,, 10

			Sleep, 200

			Send, {Enter}
			break
		}
		else if failcount < 20
		{
			Sleep, 400
			webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
			failcount++
			continue
		}
		else
		{
			Gui, Destroy
			MsgBox,, Website Error, Button didn't update with current step
			Exit
		}
	}

	try webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't get the text from the stop button after stopping
		Exit
	}

	failcount := 0
	Loop
	{
		if webstep != %currentstep%
		{
			break
		}
		else if failcount < 20
		{
			Sleep, 400
			webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
			failcount++
			continue
		}
		else
		{
			Gui, Destroy
			MsgBox,, Website Error, Button didn't update with current step
			Exit
		}
	}
	return
}

; function to retrieve patient and clinic names from the edit page
neo_getInfoFromReview()
{
    global NeoDriver, Path_ScriptNumberCSS, Path_ClinicNameCSS, Path_PatientNameCSS
	Neo_StillOpen()

	Neo_Activate()

	patientInfo := {"scriptNumber": "4324", "panNumber": "", "engravingBarcode": "", "firstName": "", "lastName": "", "clinicName": ""}

    if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")
	{
        MsgBox Not on the review page
		Exit
	}

	; get the script from the top of the review page
	try patientInfo["scriptNumber"] := NeoDriver.findElementByCss(Path_ScriptNumberCSS).Attribute("innerText")
	catch e
		MsgBox Couldn't find the script number

	patientInfo["scriptNumber"] := StrReplace(patientInfo["scriptNumber"], "Case ")

	; attempt to get the first, last, and clinic.
	try
	{
		fullname := NeoDriver.findElementByCss(Path_PatientNameCSS).Attribute("innerText")
		patientInfo["clinicName"] := NeoDriver.findElementByCss(Path_ClinicNameCSS).Attribute("innerText")
	}
	catch e
	{
		MsgBox, couldn't find patient or clinic name
		Gui, Destroy
		Exit
	}

	; take invalid characters out of the names
	invalidchars := ["...", "..", ".", ",", "'"]
	for item in invalidchars
		fullname := StrReplace(fullname, invalidchars[item], "")

	patientInfo["firstName"] := StrSplit(fullname, " ")[1]
	patientInfo["lastName"] := StrReplace(fullname, patientInfo["firstName"] " ", "")

    return patientInfo
}

; Puts new note onto the edit page
neo_newNote(orderID)
{
	global NeoDriver, Path_NewNoteCSS
	; Checks the URL to make sure it's on an edit page

	Neo_StillOpen()

	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")
	{
		MsgBox Must be on review page to use this function
		Gui, Destroy
		Exit
	}

	; New Note Button
	try NeoDriver.findElementByCss(Path_NewNoteCSS).Click()
	catch e
	{
		MsgBox,, Couldn't Find Element, Couldn't find the "new note" button
		Gui, Destroy
		Exit
	}

	BlockInput, MouseMove

	Sleep, 200

	; Dropdown for input type
	try NeoDriver.findElementByID("note_type").Click()
	catch e
	{
		BlockInput, MouseMoveOff
		MsgBox,, Couldn't Find Element, Couldn't find the note type drop down
		Gui, Destroy
		Exit
	}

	Sleep, 200

	; Select CSR
	Send {enter}

	Sleep, 300

	; enter note field

	try NeoDriver.findElementByID("note_text").Click()
	catch e
	{
		BlockInput, MouseMoveOff
		MsgBox,, Couldn't Find Element, Couldn't find the note text box
		Gui, Destroy
		Exit
	}

	Sleep, 200

	Send iTero ID: %orderID%

	BlockInput, MouseMoveOff

	return
}