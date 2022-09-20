#SingleInstance Force
CoordMode, Mouse, Client
CoordMode, Pixel, Client


f3::
{
    controlGetFocus, returnVar, Open patient case

    msgbox % returnVar

    return
}