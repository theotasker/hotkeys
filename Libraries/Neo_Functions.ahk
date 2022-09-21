; =========================================================================================================================
; RXWizard Functions
; =========================================================================================================================

global NeoCheck := false
global NeoDriver := ""
global chromeShortcutDir := A_MyDocuments "\Automation\"

ChromeGet(IP_Port := "127.0.0.1:9222") ; attaches web driver to last opened tab
{
	Driver := ComObjCreate("Selenium.ChromeDriver")
	Driver.SetCapability("debuggerAddress", IP_Port)
	Driver.Start()
	return Driver
}

neo_startWebDriver() ; closes existing Chromes and opens a new one, binds it to NeoDriver
{
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

	if !FileExist(chromeShortcutDir "ChromeForAHK.lnk")
	{
		Gui, Destroy
		MsgBox,, Chrome Shortcut Not Found, Can't find the proper shortcut at %A_MyDocuments%\Automation\ChromeForAHK.lnk `nRemember to modify the shortcut to run in debug mode by adding "--remote-debugging-port=9222"
		Exit
	}

	Run, ChromeForAHK.lnk, %chromeShortcutDir%
	Sleep, 500
	global NeoDriver := ChromeGet()
	NeoDriver.Get("https://portal.rxwizard.com/cases")

	NeoCheck := true

	currentURL := NeoDriver.Url

	return
}

neo_stillOpen() ; checks to see if a NeoDriver is bound, creates a new one if not
{
	if (NeoCheck != true)
	{
		MsgBox, 4, Open NeoDriver?, No instance of NeoDriver found, initiate?
			IfMsgBox, Yes
				neo_startWebDriver()
			IfMsgBox, No
				Exit
	}

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
	currentURL := Neo_StillOpen()

	if WinExist("New England Orthodontic Laboratory - Google Chrome")
	{
		WinActivate
	}
	else
	{
		Gui, Destroy
		MsgBox,, Website Error, RxWizard Portal should be the top tab in its own instance
		exit
	}

	if (scanField = true)
	{
		if !InStr(currentURL, "https://portal.rxwizard.com") ; Must be on a rxwizard page
		{
			Gui, Destroy
			MsgBox,, Wrong Page, Must be on an RxWizard Page
			Exit
		}  

		Send {tab}{tab}{tab}
		Sleep, 100
		NeoDriver.findElementByID("click-and-scan").click()
	}
	return currentURL
}

neo_swapPages(destPage, assignedCases:=false) ; swaps between review and edit pages
{
	currentURL := Neo_Activate(scanField=false)

	if ((destPage = "review") or (destPage = "edit") or (destPage = "swap"))
	{
		if (!InStr(currentURL, "/review/") and !InStr(currentURL, "/edit/"))
		{
			Gui, Destroy
			msgbox,, Wrong Page, Need to be on the review or edit page
			Exit
		}

		if (destPage = "review" or (destPage = "swap" and InStr(currentURL, "/edit/")))
		{
			destURL := StrReplace(currentURL, "/edit/", "/review/")
		}
		else if (destPage = "edit" or (destPage = "swap" and InStr(currentURL, "/review/")))
		{
			destURL := StrReplace(currentURL, "/review/", "/edit/")
		}
	}
	Else
	{
		destURL := "https://portal.rxwizard.com/cases"
	}
	NeoDriver.Get(destURL)

	if (assignedCases = True)
	{
		try NeoDriver.findElementByCss(casesPageCSS["assignedCases"]).Click()
		catch e
		{
			BlockInput, MouseMoveOff
			Gui, Destroy
			MsgBox,, Couldn't Find Element, Couldn't find the "Assigned to me" button
			Exit
		}

	}



	return destURL
}

neo_getPatientInfo() ; retrieves and returns patient info from review/edit pages
{
	currentURL := Neo_Activate(scanField:=false)

    if !InStr(currentURL, "https://portal.rxwizard.com/cases/review/") and !InStr(currentURL, "https://portal.rxwizard.com/cases/edit/")
	{
		Gui, Destroy
        MsgBox Not on the review or edit page
		Exit
	}

	patientInfo := {"scriptNumber": "", "panNumber": "", "engravingBarcode": "", "firstName": "", "lastName": "", "fullName": "", "clinicName": ""}

	if InStr(currentURL, "https://portal.rxwizard.com/cases/review/") ; if on review page, name is listed as full name and pan number has no ID
	{
		try patientInfo["fullName"] := NeoDriver.findElementByCss(reviewPageCSS["patientName"]).Attribute("innerText")
		catch e
		{
			Gui, Destroy
			MsgBox, couldn't find patient name
			Exit
		}

		invalidchars := ["...", "..", ".", ",", "'"] ; clean up patient name and split into first and last
		for item in invalidchars
		{
			patientInfo["fullName"] := StrReplace(patientInfo["fullName"], invalidchars[item], "")
		}
		patientInfo["firstName"] := StrSplit(patientInfo["fullName"], " ")[1]
		patientInfo["lastName"] := StrReplace(patientInfo["fullName"], patientInfo["firstName"] " ", "")

		try patientInfo["panNumber"] := NeoDriver.findElementByCss(reviewPageCSS["panNumber"], 1).Attribute("innerText")
		catch e
		{
			patientInfo["panNumber"] := "nopan"
		}

		try patientInfo["clinicName"] := NeoDriver.findElementByCss(reviewPageCSS["clinicName"]).Attribute("innerText")
		catch e
		{
			Gui, Destroy
			MsgBox, couldn't find clinic name
			Exit
		}
	}
	Else ; on edit page, patient name is broken into first and last, and pan number has an ID
	{
		try
		{
			patientInfo["firstName"] := NeoDriver.findElementByID("patient_first_name").Attribute("value")
			patientInfo["lastName"] := NeoDriver.findElementByID("patient_last_name").Attribute("value")
		}
		catch e
		{
			Gui, Destroy
			MsgBox, couldn't find patient name
			Exit
		}

		invalidchars := ["...", "..", ".", ",", "'"] ; clean up patient name and combine
		for item in invalidchars
		{
			patientInfo["firstName"] := StrReplace(patientInfo["firstName"], invalidchars[item], "")
			patientInfo["lastName"] := StrReplace(patientInfo["lastName"], invalidchars[item], "")
		}
		patientInfo["fullName"] := patientInfo["firstName"] " " patientInfo["lastName"]

		try patientInfo["panNumber"] := NeoDriver.findElementByID("pan").Attribute("innerText")
		catch e
		{
			patientInfo["panNumber"] := "nopan"
		}

		try patientInfo["clinicName"] := NeoDriver.findElementByID("office").Attribute("innerText")
		catch e
		{
			Gui, Destroy
			MsgBox, couldn't find clinic name
			Exit
		}
	}

	try patientInfo["scriptNumber"] := NeoDriver.findElementByClass("case-name").Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox Couldn't find the script number
	}

	patientInfo["panNumber"] := StrReplace(patientInfo["panNumber"], " ", "") ; clean up all gathered patient info
	patientInfo["scriptNumber"] := StrReplace(patientInfo["scriptNumber"], "Case ")
	patientInitials := SubStr(patientInfo["firstName"], 1, 1) SubStr(patientInfo["lastName"], 1, 1)
	patientInfo["engravingBarcode"] := StrSplit(patientInfo["scriptNumber"], "-")[2] "-" patientInitials "-" patientInfo["panNumber"]
	patientInfo["clinicName"] := StrReplace(patientInfo["clinicName"], "#", "")
	patientInfo["clinicName"] := StrReplace(patientInfo["clinicName"], ".", "")

    return patientInfo
}

neo_newNote(orderID) ; Puts new note onto the edit page
{
	global NeoDriver
	BlockInput MouseMove

	Neo_StillOpen()

	if (!InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/")) and (!InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/"))
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox Must be on review or edit page to use this function
		Exit
	}

	try NeoDriver.findElementByCss(bothPageCSS["newNote"]).Click()
	catch e
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Couldn't Find Element, Couldn't find the "new note" button
		Exit
	}
	Sleep, 200

	try NeoDriver.findElementByID("note_type").Click() ; Dropdown for input type
	catch e
	{
		BlockInput, MouseMoveOff
		Gui, Destroy
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
		Gui, Destroy
		MsgBox,, Couldn't Find Element, Couldn't find the note text box
		Exit
	}

	Sleep, 200
	Send iTero ID: %orderID%

	try NeoDriver.findElementByCss(bothPageCSS["noteSave"]).Click()
	catch e
	{
		BlockInput, MouseMoveOff
		Gui, Destroy
		MsgBox,, Couldn't Find Element, Couldn't find the note save button
		Gui, Destroy
		Exit
	}

	BlockInput, MouseMoveOff
	return
}

neo_uploadPic(screenshotDir) {
	global NeoDriver
	BlockInput MouseMove

	try NeoDriver.findElementByCss(bothPageCSS["uploadFile"]).Click()
	catch e 
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Website Error, Couldn't find the file upload button on the website, make sure case is in production
		Exit
	}

	WinWaitActive Open,, 5
	if ErrorLevel 
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Website Error, File Upload window didn't open properly
		Exit
	}

	Sleep, 100

	Send % screenshotDir ; send the directory for the automation screencaps folder
	Sleep, 200
	Send {enter}
	Sleep, 200

	Send {shiftDown}{tab}{ShiftUp} ; tab into the main files window
	Sleep, 100

	Send {CtrlDown}a{CtrlUp} ; select all files
	Sleep, 100

	Send {enter} ; confirm

	BlockInput, MouseMoveOff
	return
}

; =========================================================================================================================
; Deprecated, updates to RXWizard made obsolete
; =========================================================================================================================

neo_start(currentStep) ; hits the start button on the review page
{
	global NeoDriver

	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/") ; checks to ensure on the review page
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Wrong Page, Must be on the review page for this function
		Exit
	}

	; Wait until the button for pushing step is clickable
	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.CSS_SELECTOR, currentstepxpath))
	catch e
	{
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the start button before it was clicked the first time
		Exit
	}

	; Try to find the text label on the button for pushing steps
	try webstep := NeoDriver.findElementByCss(currentstepxpath).Attribute("innerText")
	catch e
	{
		BlockInput MouseMoveOff
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
			BlockInput MouseMoveOff
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
		BlockInput MouseMoveOff
		Gui, Destroy
		MsgBox,, Web Error, Button wasn't on the right step
		Exit
	}
	return
}