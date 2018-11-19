/*

Writes a timer script to Startup-directory. The Wake-Up-Timer script deletes itself when the timer is finished
Hibernates the computer, depending on variable "Hibernate"

*/


; *************
; CONFIGURATION
; *************

;run, autohotkey.exe %WakeUpMain% %YYYY% %MM% %DD% %Hour% %Min% %Standby% %Start% %Resume% %alldays%

YYYY=%1%				;Parameter 1: Year  
MM=%2%					;Parameter 2: Month
DD=%3%					;Parameter 3: Day
Hour=%4%				;Parameter 4: hour
Min=%5%					;Parameter 5: minute

Hibernate=%6%			;Parameter 6: Hibernate, suspend, keep running?
Start=%7%				;Parameter 7: Application to run
Resume=%8%				;Parameter 8; Resume to run application? Yes/No

alldays=%9%				;Parameter 9; Repeat every days? Yes/No

ScheduledTime=%YYYY%%MM%%DD%%Hour%%Min%%A_sec%

; ***********
; AUTOEXECUTE
; ***********

;Get name of application to be started from it's path:
SplitPath, Start, FileName 
StringReplace, FileName,FileName,", ;Remove "

;Writes the timer script to Startup-directory
WriteAutostartFile(YYYY,MM,DD,Hour,Min,Start,Resume,FileName,alldays)

;Starts the timer from Startup-directory
Run, %A_startup%\%ScheduledTime%_%Filename%.ahk

;Hibernates the computer, depending on variable "Hibernate": 
	If Hibernate=1 		;Hibernate
			{ 
			DllCall("PowrProf\SetSuspendState", "int", 1, "int", 0, "int", 0) 
			} 

	If Hibernate=2		;Suspend
			{
			DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0) 
			}


; *********
; FUNCTIONS
; *********

WriteAutostartFile(YYYY,MM,DD,Hour,Min,Start,Resume,FileName,alldays)
; Writes Wake-Up-Timer script to Startup directory . After a restart the script will run again and set the timer. 
; The Wake-Up-Timer script deletes itself when the timer is finished or when duetime is over at startup
{

    if (alldays=1)
    {
         ; run every days

        ScheduledTime=%YYYY%%MM%%DD%%Hour%%Min%%A_sec%
        FileAppend,
        (
        #Persistent
        Menu Tray, Icon, shell32.dll, 240
        Menu, tray, tip,  Wake-Up-Timer``nStart Time: %Hour%:%Min%``nApplication: %FileName%

        Hour=%Hour%				
        Min=%Min%				
        Start=%Start%				; Application to run
        Resume=%Resume%				; Resume to run application? Yes/No
        

        SetTimer, CheckTime, 10000	; Time in miliseconds (1000 = 1s)
        Return

        CheckTime:	
            myTime := A_Hour A_Min
            Today := A_MDAY

            If (myTime = %Hour%%Min% && Today <> Ran)
            {
            run, %Start%
            Ran := A_MDAY
            }

        Return	
        ), %A_startup%\%ScheduledTime%_%Filename%.ahk
        
    }
    else
    {
        ;single run
        ScheduledTime=%YYYY%%MM%%DD%%Hour%%Min%%A_sec%
        FileAppend,
        (
        Menu Tray, Icon, shell32.dll, 240

        YYYY=%YYYY%				
        MM=%MM%					
        DD=%DD%					
        Hour=%Hour%				
        Min=%Min%				
        Start=%Start%				; Application to run
        Resume=%Resume%				; Resume to run application? Yes/No
        
        ScheduledTime=%YYYY%%MM%%DD%%Hour%%Min%%A_sec%
        Name=%A_Now%				;Name of the Timer Object
        
        If (A_Now>ScheduledTime)
        {
                ;x := (A_Now-%ScheduledTime%)
    ;			msgbox, `%A_Now`%`(A_Now)``n%ScheduledTime% (scheduledTime)``n`%x`%
    ;			exitApp
            msgbox, The application ``n     %FileName% ``ncould not be started as scheduled``n     %YYYY%-%MM%-%DD% %Hour%:%Min%
            FileDelete, %A_startup%\%ScheduledTime%_%Filename%.ahk
            exitApp
        }
        else
        {
            Menu, tray, tip,  Wake-Up-Timer``nStart Time: %YYYY%-%MM%-%DD% %Hour%:%Min%``nApplication: %FileName%
            WakeUp(YYYY, MM, DD, Hour, Min, Resume, Name) 
            run, %Start%
            FileDelete, %A_startup%\%ScheduledTime%_%Filename%.ahk
            return
        }	

        ; FUNCTIONS 
        WakeUp(Year, Month, Day, Hour, Minute, Resume, Name) 
        ;Awaits duetime, then returns to the caller (like some sort of "sleep until duetime"). 
        ;If the computer is in hibernate or suspend mode 
        ;at duetime, it will be reactivated (hardware support provided) 
        ;Parameters: Year, Month, Day, Hour, Minute together produce duetime 
        ;Resume: If Resume=1, the system is restored from power save mode at due time
        ;Name: Arbitrary name for the timer
        { 
            duetime:=GetUTCFileTime(Year, Month, Day, Hour, Minute) 

            Handle:=DLLCall("CreateWaitableTimer" 
                    ,"char *", 0 
                    ,"Int",0 
                    ,"Str",name, "UInt") 

            DLLCall("CancelWaitableTimer","UInt",handle) 

            DLLCall("SetWaitableTimer" 
                  ,"Uint", handle 
                  ,"Int64*", duetime        ;duetime must be in UTC-file-time format! 
                  ,"Int", 1000 
                  ,"uint",0 
                  ,"uint",0 
                  ,"int",resume) 



            Signal:=DLLCall("WaitForSingleObject" 
                    ,"Uint", handle 
                    ,"Uint",-1) 

            DllCall("CloseHandle", uint, Handle)   ;Closes the handle
        }


        GetUTCFiletime(Year, Month, Day, Hour, Min) 
        ;Converts "System Time" (readable time format) to "UTC File Time" (number of 100-nanosecond intervals since January 1, 1601 in  Coordinated Universal Time UTC) 
        { 
            DayOfWeek=0 

            Second=00 
            Millisecond=00 


            ;Converts System Time to Local File Time: 
            VarSetCapacity(MyFiletime  , 64, 0) 
            VarSetCapacity(MySystemtime, 32, 0) 

            InsertInteger(Year,       MySystemtime,0) 
            InsertInteger(Month,      MySystemtime,2) 
            InsertInteger(DayOfWeek,  MySystemtime,4) 
            InsertInteger(Day,        MySystemtime,6) 
            InsertInteger(Hour,       MySystemtime,8) 
            InsertInteger(Min,        MySystemtime,10) 
            InsertInteger(Second,     MySystemtime,12) 
            InsertInteger(Millisecond,MySystemtime,14) 

            DllCall("SystemTimeToFileTime", Str, MySystemtime, UInt, &MyFiletime) 
            LocalFiletime := ExtractInteger(MyFiletime, 0, false, 8) 

            ;Converts local file time to a file time based on the Coordinated Universal Time (UTC): 
            VarSetCapacity(MyUTCFiletime  , 64, 0) 
            DllCall("LocalFileTimeToFileTime", Str, MyFiletime, UInt, &MyUTCFiletime) 
            UTCFiletime := ExtractInteger(MyUTCFiletime, 0, false, 8) 

            Return UTCFileTime 
        } 


        ExtractInteger(ByRef pSource, pOffset = 0, pIsSigned = false, pSize = 32) 
        ; Documented in Autohotkey Help 
        { 
            Loop `%pSize`%  
                result += *(&pSource + pOffset + A_Index-1) << 8*(A_Index-1) 
            if (!pIsSigned OR pSize > 4 OR result < 0x80000000) 
                return result  
            return -(0xFFFFFFFF - result + 1) 
        } 


        InsertInteger(pInteger, ByRef pDest, pOffset = 0, pSize = 4) 
        ; Documentated in Autohotkey Help 
        { 
          Loop `%pSize`% 
                  DllCall("RtlFillMemory", UInt, &pDest + pOffset + A_Index-1 
                          , UInt, 1, UChar, pInteger >> 8*(A_Index-1) & 0xFF) 
        }
        
        ), %A_startup%\%ScheduledTime%_%Filename%.ahk
    }
}