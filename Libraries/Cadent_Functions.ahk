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
	global MyCadentDriver, Path_CadentSearchFieldID
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
    MyCadentDriver.findElementByID(Path_CadentSearchFieldID).click()
	Sleep, 100
	Send, {CtrlDown}a{CtrlUp}

	return
}

Cadent_GetOrderID()
{
	global MyCadentDriver, Path_CadentOrderNumberID

	Cadent_StillOpen()

	if !InStr(MyCadentDriver.Url, "https://mycadent.com/CaseInfo.aspx")
	{
		MsgBox Must be on MyCadent order page to use this function
		Gui, Destroy
		Exit
	}

	orderID := MyCadentDriver.findElementByID(Path_CadentOrderNumberID).Attribute("value")

	return orderID

}