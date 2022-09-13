; Library for all standalone functions pertaining to the operation of Chrome through Selenium

; Don't Touch, attaches web driver to last opened tab
ChromeGet(IP_Port := "127.0.0.1:9222") 
	{
		Driver := ComObjCreate("Selenium.ChromeDriver")
		Driver.SetCapability("debuggerAddress", IP_Port)
		Driver.Start()
		return Driver
	}

; ================================================================================================================
; RXWizard functions
; ================================================================================================================

Neo_StartWebDriver()
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
	
	NeoCheck := 1
	
	Exit
}

; function to check that the NeoDriver is still open
Neo_StillOpen()
{
	global NeoDriver, NeoCheck
	
	if NeoCheck = 0
		MsgBox, 4, Open NeoDriver?, No instance of NeoDriver found, initiate?
			IfMsgBox, Yes
				Neo_StartWebDriver()
			IfMsgBox, No
				Exit
	
	try temp := InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	catch e
	{
		MsgBox, 4, Webdriver Error, The tab for driving Portal.RXWizard was closed, initiate webdriver?
			IfMsgBox, Yes
				Neo_StartWebDriver()
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
Neo_Activate()
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
Neo_NavigateToCases()
{
    global NeoDriver
	Neo_StillOpen()

    if (NeoDriver.Url != "https://portal.rxwizard.com/cases") 
    {
        NeoDriver.Get("https://portal.rxwizard.com/cases")
    }
	
	Neo_Activate()
	
	Send {tab}
	Sleep, 100
    NeoDriver.findElementByCss(".ant-input:nth-child(2)").click()

}

; function for getting to the review page from the edit page
Neo_NavigateReviewFromEdit()
{
    global NeoDriver
	Neo_StillOpen()
	
	Neo_Activate()
	
	; If on the edit page, click on the "review" button on the bottom, then either wait for the review page or
	; click the "confirm" button on the "over model count" popup.
    if InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	{
        try NeoDriver.findElementByCss("#main-content-wrapper > div > div.content.ant-layout-content > form > div.ant-row-flex.ant-row-flex-end > div:nth-child(2) > button").Click()
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
Neo_Start2()
{
	global NeoDriver, currentstep, currentstepxpath
		
	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/review/") ; checks to ensure on the review page
	{
		MsgBox,, Wrong Page, Must be on the review page for this function
		Gui, Destroy
		Exit
	}
	
	; Wait until the button for pushing step is clickable
	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.XPATH, currentstepxpath))
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the start button before it was clicked the first time
		Exit
	}
		
	; Try to find the text label on the button for pushing steps
	try webstep := NeoDriver.findElementByXpath(currentstepxpath).Attribute("innerText")
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't get the text from the start stop button before starting
		Exit
	}
	
	; check the retrieved button text against the currenstep variable
	if webstep = %currentstep%
	{
		try NeoDriver.findElementByXpath(currentstepxpath).sendKeys(Keys.RETURN)
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
Neo_Stop2()
{
	global NeoDriver, currentstep, currentstepxpath
	
	; Wait until the button for pushing step is clickable
	try new WebDriverWait(NeoDriver, 10).until(ExpectedConditions.element_to_be_clickable(By.XPATH, currentstepxpath))
	catch e
	{
		Gui, Destroy
		MsgBox,, Web Error, Couldn't find the stop button before it was clicked 
		Exit
	}
	
	; Try to find the text label on the button for pushing steps
	try webstep := NeoDriver.findElementByXpath(currentstepxpath).Attribute("innerText")
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
			NeoDriver.findElementByXpath(currentstepxpath).sendKeys(Keys.RETURN)
						
			Neo_Activate()
			
			WinWaitActive, New England Orthodontic Laboratory - Google Chrome,, 10
			
			Sleep, 200
			
			Send, {Enter}
			break
		}
		else if failcount < 20
		{
			Sleep, 400
			webstep := NeoDriver.findElementByXpath(currentstepxpath).Attribute("innerText")
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
	
	try webstep := NeoDriver.findElementByXpath(currentstepxpath).Attribute("innerText")
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
			webstep := NeoDriver.findElementByXpath(currentstepxpath).Attribute("innerText")
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
Neo_GetInfoFromEdit()
{
    global NeoDriver, firstname, lastname, clinic, scriptnumber
	Neo_StillOpen()
	
	Neo_Activate()
	
    if InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
    {
        ; get the script from the top of the review page
        try scriptnumber := NeoDriver.findElementByCss("#main-content-wrapper > div > div.header > div.name > div > span.case-name").Attribute("innerText")
		catch e
			MsgBox Couldn't find the script number
		
        scriptnumber := StrReplace(scriptnumber, "Case ")
        
		; attempt to get the first, last, and clinic.
        try
		{
			firstname := NeoDriver.findElementByID("patient_first_name").Attribute("value")
			lastname := NeoDriver.findElementByID("patient_last_name").Attribute("value")
			clinic := NeoDriver.findElementByCss("#office > div > div > div.ant-select-selection-selected-value").Attribute("title")
		}
		catch e
		{
			MsgBox, Couldn't find one of the variables, ending function
			Gui, Destroy
			Exit
		}
		
		; take invalid characters out of the names
		invalidchars := ["...", "..", ".", ",", "'"]

		for item in invalidchars
			firstname := StrReplace(firstname, invalidchars[item], "")
		
		for item in invalidchars
			lastname := StrReplace(lastname, invalidchars[item], "")
    }
    else
	{
        MsgBox Not on the edit page
		Exit
	}
    return
}

; Puts new note onto the edit page
Neo_NewNote()
{
	global NeoDriver, orderid
	; Checks the URL to make sure it's on an edit page

	Neo_StillOpen()

	if !InStr(NeoDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	{
		MsgBox Must be on edit page to use this function
		Gui, Destroy
		Exit
	}
	
	; New Note Button
	try NeoDriver.findElementByCss("#main-content-wrapper > div > div.content.ant-layout-content > form > div.ant-row > div:nth-child(2) > div:nth-child(3) > div > button").Click()
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
	
	return
}
	
; ================================================================================================================
; MyCadent Functions
; ================================================================================================================

Cadent_StartWebDriver()
{
	global
	
	; Check to make sure that the modified Chrome shortcut is in the right place
	chromelocation = %A_MyDocuments%\Automation\ChromeForAHK.lnk
	if !FileExist(chromelocation)
	{
		MsgBox,, Chrome Shortcut Not Found, Can't find the proper shortcut at %A_MyDocuments%\Automation\ChromeForAHK.lnk `nRemember to modify the shortcut to run in debug mode by adding "--remote-debugging-port=9222"
		Exit
	}
	
	; Open another window and assign it to mycadent
	Run, ChromeForAHK.lnk, %A_MyDocuments%\Automation\
	Sleep, 500
	MyCadentDriver := ChromeGet()
	MyCadentDriver.Get("https://mycadent.com/COrdersList.aspx")
	
	CadentCheck := 1
	
	Gui, Destroy
	Exit
}

; function to check that Cadent is still open
Cadent_StillOpen()
{
	global MyCadentDriver, CadentCheck
	
	if CadentCheck = 0
	{
		Gui, Destroy
		MsgBox, 4, Open CadentDriver?, No instance of CadentDriver found, initiate?
			IfMsgBox, Yes
			{
				Cadent_StartWebDriver()
			}
			IfMsgBox, No
			{
				Gui, Destroy
				Exit
			}
	}
	
	try temp := InStr(MyCadentDriver.Url, "https://portal.rxwizard.com/cases/edit/")
	catch e
	{
		MsgBox, 4, Webdriver Error, The tab for driving MyCadent was closed, reinitiate webdriver?
			IfMsgBox, Yes
				Cadent_StartWebDriver()
			IfMsgBox, No
			{
				return
			}
		Gui, Destroy		
		Exit
	}
}

; function to go to orders page
Cadent_Orders()
{
	global MyCadentDriver
	Cadent_StillOpen()

	; Check the current page of the cadent driver
	if (MyCadentDriver.Url != "https://mycadent.com/COrdersList.aspx") 
    {
        MyCadentDriver.Get("https://mycadent.com/COrdersList.aspx")
    }
	
	if WinExist("Clinical Orders List - Google Chrome")
		WinActivate
	else
	{
		MsgBox The MyCadent site must be on the top tab for this shortcut to work
		Gui, Destroy
		Exit
	}
	
	Sleep, 100
	Send, {tab}
	Sleep, 100
    MyCadentDriver.findElementByID("ctl00_body_OrdersListReport_ctl01_ctl05_ctl05-string-operand").click()
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}
	
	return
}
	
Cadent_GetOrderID()
{
	global MyCadentDriver, orderid
	
	Cadent_StillOpen()
	
	if !InStr(MyCadentDriver.Url, "https://mycadent.com/CaseInfo.aspx")
	{
		MsgBox Must be on MyCadent order page to use this function
		Gui, Destroy
		Exit
	}
	
	orderid := MyCadentDriver.findElementByID("ctl00_body_txtOrderHeaderID").Attribute("value")
	
	return
	
}