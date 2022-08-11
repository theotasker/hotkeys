#SingleInstance, Force

ABR_Output_Dir := "\\app03\Scans\~Digital Dept Share\R&D Network\EasyRX\Auto Import Output\"
Input_Dir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Input\"
Error_Dir := "\\app03\Scans\~Digital Dept Share\R&D Network\Auto Importing\Errors\"

Loop
{
	if FileExist(ABR_Output_Dir "*.stl")  ; Main Loop
	{

		Loop files, %ABR_Output_Dir%*.stl  
		{
			Temp_File := A_LoopFileName
			If InStr(Temp_File, "[1]") ; single arch expected, move right away
			{
				FileMove, %ABR_Output_Dir%%Temp_File%, %Input_Dir%%Temp_File%, 1
			}
			
			Else if InStr(Temp_File, "[2]") ; Two arches expected, check for opposing arch before moving
			{
				if InStr(Temp_File, "~Upr") ; If it's an upper, look for the lower before moving
				{
					Lower_File := StrReplace(Temp_File, "~Upr", "~Lwr")
					if FileExist(ABR_Output_Dir Lower_File)
					{
						FileMove, %ABR_Output_Dir%%Temp_File%, %Input_Dir%%Temp_File%, 1
						FileMove, %ABR_Output_Dir%%Lower_File%, %Input_Dir%%Lower_File%, 1
					}
					Else
					{
						Sleep, 10000 ;  Wait for lower
					}
				}

				Else if InStr(Temp_File, "~Lwr") ; If it's a Lower, look for the upper before moving
				{
					Upper_File := StrReplace(Temp_File, "~Lwr", "~Upr")
					if FileExist(ABR_Output_Dir Upper_File)
					{
						FileMove, %ABR_Output_Dir%%Temp_File%, %Input_Dir%%Temp_File%, 1
						FileMove, %ABR_Output_Dir%%Upper_File%, %Input_Dir%%Upper_File%, 1
					}
					Else
					{
						Sleep, 10000 ;  Wait for upper
					}
				}
				
			}
			else
			{
				FileMove, ABR_Output_Dir Temp_File, Error_Dir Temp_File, 1
			}

		}
	}
	
	else
	{
		Sleep, 10000
	}
	
}