#SingleInstance Force
CoordMode, Mouse, Client
CoordMode, Pixel, Client

#NoEnv

currentURL := "https://portal.rxwizard.com/cases/review/486639"


if (!InStr(currentURL, "/review/") and !InStr(currentURL, "/edit/"))
{
    msgbox % currentURL
    msgbox,, Wrong Page, Need to be on the review or edit page
    Exit
}
Else
{
    msgbox all good
}