#SingleInstance Force
CoordMode, Mouse, Client
CoordMode, Pixel, Client

SetTitleMatchMode, 2

f3::
{
    ControlGetText, activeTool, TdfGroupInfo1, Feature
    msgbox % activeTool

    return
}