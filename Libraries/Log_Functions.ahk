; ===========================================================================================================================
; Library for all functions pertaining to Logging
; ===========================================================================================================================

; start variables for logging numbers for the server txt, will reset to zero every time they're logged
importedserver = 0
preppedserver = 0
engravedserver = 0
hotkeysserver = 0
clicksserver = 0
strokesserver = 0

; start variables for logging numbers for the local txt, will continue on counting after importing from local log
importedlocal = 0
preppedlocal = 0
engravedlocal = 0
hotkeyslocal = 0
clickslocal = 0
strokeslocal = 0

Log_ImportLocalLog() ; Imports the current log count from the local log, creates a log if none exists
{
   global
   LocalLogDir := A_MyDocuments "\Automation\Logs\" ;folder for logs
   LocalLogFile := A_MyDocuments "\Automation\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" ; log file, named with todays date

   ; sets the text in case it has to create a new txt file
   LocalLogText := Log_GenerateText(importedlocal, preppedlocal, engravedlocal, hotkeyslocal, clickslocal, strokeslocal)

   if !FileExist(LocalLogDir)
   {
      FileCreateDir, %LocalLogDir%
   }
   if !FileExist(LocalLogFile)
   {
      FileAppend, %LocalLogText%, %LocalLogfile% ; creates a new log file with the current variables
   }

   FileRead, ExtantLog, %LocalLogfile%

   LogArray := StrSplit(ExtantLog, "`n") ; creates an array of the individual lines of the txt

   ; for each variable, remove the text and convert the number string to a integer
   importedlocal := GetNumber(LogArray[1], "Cases imported = ")
   preppedlocal := GetNumber(LogArray[2], "Cases prepped = ")
   engravedlocal := GetNumber(LogArray[3], "Cases engraved = ")
   hotkeyslocal := GetNumber(LogArray[4], "Hotkeys used = ")
   clickslocal := GetNumber(LogArray[5], "Mouse actions saved = ")
   strokeslocal := GetNumber(LogArray[6], "Keystrokes saved = ")

   return
}

; function to be called in main program, increments the log by (#clicks, #strokes), and 1 hotkeysused
Log_Increment(clicks = 0, strokes = 0)
{
   global

   clickslocal := clicks + clickslocal
   clicksserver := clicks + clicksserver

   strokeslocal := strokes + strokeslocal
   strokesserver := strokes + strokesserver

   hotkeyslocal+=1
   hotkeysserver+=1

   return
}

; updates the log txt with the current log variables plus existing txt variables, then sets local variables to 0
Log_UpdateServer()
{
   global

   ServerLogDir := "\\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Logs\" ;folder for logs
   ServerLogFile := "\\APP03\Scans\~Digital Dept Share\R&D Network\Scripting\AutoHotKey\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt"

   ; sets the text in case it has to create a new txt file
   ServerLogText := Log_GenerateText(importedserver, preppedserver, engravedserver, hotkeysserver, clicksserver, strokesserver)

   if !FileExist(ServerLogDir)
   {
      FileCreateDir, %ServerLogDir%
   }
   if !FileExist(ServerLogFile)
   {
      FileAppend, %ServerLogText%, %ServerLogfile% ; creates a new log file with the current variables
   }

   FileRead, ExtantLog, %ServerLogfile%

   LogArray := StrSplit(ExtantLog, "`n") ; creates an array of the individual lines of the txt

   ; for each variable, remove the text and convert the number string to a integer
   importadd := GetNumber(LogArray[1], "Cases imported = ")
   prepadd := GetNumber(LogArray[2], "Cases prepped = ")
   engraveadd := GetNumber(LogArray[3], "Cases engraved = ")
   hotkeysadd := GetNumber(LogArray[4], "Hotkeys used = ")
   clicksadd := GetNumber(LogArray[5], "Mouse actions saved = ")
   keystrokesadd := GetNumber(LogArray[6], "Keystrokes saved = ")

   ; adds 1 to the corresponding current step
   if Step = Importing
   {
      importedserver = 1
   }
   else if Step = Prepping
   {
      preppedserver = 1
   }
   else if Step = Engraving
   {
      engravedserver = 1
   }

   ; adds the local log variables with the corresponding variables from the txt
   importedserver := importedserver + importadd
   preppedserver := preppedserver + prepadd
   engravedserver := engravedserver + engraveadd
   hotkeysserver := hotkeysserver + hotkeysadd
   clicksserver := clicksserver + clicksadd
   strokesserver := strokesserver + keystrokesadd

   ; updates the log text, then writes to the file
   ServerLogText := Log_GenerateText(importedserver, preppedserver, engravedserver, hotkeysserver, clicksserver, strokesserver)
   FileDelete, %ServerLogFile%
   FileAppend, %ServerLogText%, %ServerLogfile%

   ; reset server variables to 0
   importedserver = 0
   preppedserver = 0
   engravedserver = 0
   hotkeysserver = 0
   clicksserver = 0
   strokesserver = 0

   return
}

Log_UpdateLocal()
{
   global

   ; just in case it needs to update to todays date
   LocalLogFile := A_MyDocuments "\Automation\Logs\" A_MM "-" A_DD "-" A_YYYY ".txt" ; log file, named with todays date

   if Step = Importing
   {
      importedlocal += 1
   }
   else if Step = Prepping
   {
      preppedlocal += 1
   }
   else if Step = Engraving
   {
      engravedlocal += 1
   }

   LocalLogText := Log_GenerateText(importedlocal, preppedlocal, engravedlocal, hotkeyslocal, clickslocal, strokeslocal)
   FileDelete, %LocalLogFile%
   FileAppend, %LocalLogText%, %LocalLogfile%
   return
}

; generated the text to be put into the txt file, based on the variables passed
Log_GenerateText(imported = 0, prepped = 0, engraved = 0, hotkeys = 0, clicks = 0, strokes = 0)
{
   LogText = Cases imported = %imported%`nCases prepped = %prepped%`nCases engraved = %engraved%`nHotkeys used = %hotkeys%`nMouse actions saved = %clicks%`nKeystrokes saved = %strokes%
   return LogText
}

; function to get the number from the individual string from the txt log
GetNumber(x = "", y = "")
{
   z := StrReplace(x, y, "") ; removes the leading text
   z := StrReplace(z, "`n", "") ; removes the new line markers
   z += 0 ; makes the number an integer
   return z
}