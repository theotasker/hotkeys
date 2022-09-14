#SingleInstance Force
CoordMode, Mouse, Client
CoordMode, Pixel, Client

currentURL := "https://portal.rxwizard.com/cases/edit/486689"


testMe(destPage)
{
    msgbox % destPage

    if (destPage = "review" or (destPage = "swap" and InStr(currentURL, "/edit/")))
    {
        destURL = StrReplace(currentURL, "/edit/", "/review/")
    }
    else if (destPage = "edit" or (destPage = "swap" and InStr(currentURL, "/review/")))
    {
        destURL = StrReplace(currentURL, "/review/", "/edit/")
    }
    Else
    {
        msgbox faaail 
        exit
    }

}

testMe(destPage:="review")
