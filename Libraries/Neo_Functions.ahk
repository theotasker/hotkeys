NeoCheck := false

neo_startWebDriver() ; closes existing Chromes and opens a new one, binds it to NeoDriver
{
	global

	if WinExist("ahk_exe chrome.exe")
		MsgBox, 4, Close Chrome?, Close all current instances of Chrome? (suggested)
			IfMsgBox, Yes
			{
				While WinExist("ahk_exe chrome.exe") ; close all instances of chrome
				{
					Loop, 10
					{
						WinClose, ahk_exe chrome.exe
					}
				}
			}
			IfMsgBox, No
			{
				return
			}

	chromelocation = %A_MyDocuments%\Automation\ChromeForAHK.lnk
	if !FileExist(chromelocation)
	{
		MsgBox,, Chrome Shortcut Not Found, Can't find the proper shortcut at %A_MyDocuments%\Automation\ChromeForAHK.lnk `nRemember to modify the shortcut to run in debug mode by adding "--remote-debugging-port=9222"
		Exit
	}

	Run, ChromeForAHK.lnk, %A_MyDocuments%\Automation\
	Sleep, 500
	NeoDriver := ChromeGet()
	NeoDriver.Get("https://portal.rxwizard.com/cases")

	NeoCheck := true

	return NeoDriver
}


neo_stillOpen() ; checks to see if a NeoDriver is bound, creates a new one if not
{
	global NeoDriver, NeoCheck

	if (NeoCheck != true)
		MsgBox, 4, Open NeoDriver?, No instance of NeoDriver found, initiate?
			IfMsgBox, Yes
				neo_startWebDriver()
			IfMsgBox, No
				Exit

	try currentURL := NeoDriver.Url

	catch e
	{
		MsgBox, 4, Webdriver Error, The tab for driving Portal.RXWizard was closed, initiate webdriver?
			IfMsgBox, Yes
				neo_startWebDriver()
			IfMsgBox, No
			{
				Exit
			}
		return "https://portal.rxwizard.com/cases"
	}
	return currentURL
}


neo_activate(scanField) ; Bring the Chrome running RXWizard to the front, pop into scan field if requested
{
	global NeoDriver, Path_ScanScriptCSS

	if WinExist("New England Orthodontic Laboratory - Google Chrome")
		WinActivate
	else
	{
		MsgBox,, Website Error, RxWizard Portal should be the top tab in its own instance
		exit
	}

	if scanField
	{
		Neo_StillOpen()

		if !InStr(NeoDriver.Url, "https://portal.rxwizard.com") ; Must be on a rxwizard page
		{
			Gui, Destroy
			MsgBox,, Wrong Page, Must be on an RxWizard Page
			Exit
		}

		Send {tab}
		Sleep, 100
		NeoDriver.findElementByCss(Path_ScanScriptCSS).click()
	}
	return
}


neo_swapPages(destPage) ; swaps between review and edit pages
{
    global NeoDriver, Path_ReviewButtonCSS

	currentURL := Neo_StillOpen()

	Neo_Activate(scanField=false)

	if (!InStr(currentURL, "/review/") and !InStr(currentURL, "/edit/"))
	{
		msgbox,, Wrong Page, Need to be on the review or edit page
		Exit
	}

	if (destPage = "review" or (destPage = "swap" and InStr(currentURL, "/edit/")))
	{
		destURL = StrReplace(currentURL, "/edit/", "/review/")
	}
	else if (destPage = "edit" or (destPage = "swap" and InStr(currentURL, "/review/")))
	{
		destURL = StrReplace(currentURL, "/review/", "/edit/")
	}
	else

	NeoDriver.Get(destURL)
	return destURL
}

neo_start() ; hits the start button on the review page
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

		Neo_Activate(scanField=false)

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

neo_Stop() ; hits the stop button on the review page
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

			Neo_Activate(scanField=false)

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

neo_getInfoFromReview() ; retrieves and returns patient info from review/edit pages
{
    global NeoDriver, Path_ScriptNumberCSS, Path_ClinicNameCSS, Path_PatientNameCSS
	Neo_StillOpen()

	Neo_Activate(scanField=false)

	patientInfo := {"scriptNumber": "", "panNumber": "", "engravingBarcode": "", "firstName": "", "lastName": "", "clinicName": ""}

    if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/") and !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	{
        MsgBox Not on the review or edit page
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

	Neo_StillOpen()

	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")
	{
		MsgBox Must be on review page to use this function
		Exit
	}

	try NeoDriver.findElementByCss(Path_NewNoteCSS).Click()
	catch e
	{
		MsgBox,, Couldn't Find Element, Couldn't find the "new note" button
		Exit
	}

	BlockInput, MouseMove
	Sleep, 200

	try NeoDriver.findElementByID("note_type").Click() ; Dropdown for input type
	catch e
	{
		BlockInput, MouseMoveOff
		MsgBox,, Couldn't Find Element, Couldn't find the note type drop down
		Exit
	}

	Sleep, 200
	Send {enter} ; Select CSR
	Sleep, 300

	try NeoDriver.findElementByID("note_text").Click() ; enter note field
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