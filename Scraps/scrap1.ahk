#SingleInstance Force
CoordMode, Mouse, Client
CoordMode, Pixel, Client

global testMe := {"":""
, "greatings": "rightous"
, "reawdwa":"dwijdijwdj"}

running()
{
    msgbox % testMe["greatings"]
}

running()