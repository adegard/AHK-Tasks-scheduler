/*
@adegard, 2018 AHK-Tasks-scheduler-v1.ahk   https://github.com/adegard/AHK-Tasks-scheduler
1-How to install: PutWapUp.ahk file in the same folder
2-Run AHK-Task-scheduler and choose a file to schedule:
    -schedule single task: job is deleted after run
    -schedule every task: job remains minimized in tray and start each day (Every day check)

adapted from original: https://autohotkey.com/board/topic/11200-ahk-scheduler/
*/

WakeUpMain=WakeUp.ahk

Menu Tray, Icon, shell32.dll, 240
Gui Add, GroupBox, x8 y8 w251 h183, Schedule a new Task:
Gui Add, Text, x24 y96 w60 h20, StartDate:
Gui Add, Text, x24 y128 w60 h20, StartTime:
Gui Add, DateTime, x96 y96 w80 h20 vStartDate Section, 
Gui Add, DateTime, x96 y128 w80 h20 vStartTime 1, HH:mm:00 ;time 
Gui Add, Text, x24 y32 w80 h20, Run (with path):
Gui Add, GroupBox, x8 y200 w250 h66, Wake up from Standby?
Gui Add, Radio, vWakeUp x24 y232 w36 h13, &No
Gui Add, Radio, x72 y232 w40 h13 +Checked, &Yes
Gui Add, GroupBox, x272 y8 w150 h126, Power Management
Gui Add, Radio, x304 y32 w80 h26 +Checked, &Keep running
Gui Add, Radio, vStandby x304 y64 w80 h26, &Hibernate NOW
Gui Add, Radio, x304 y104 w80 h26, &Suspend NOW
Gui Add, Button, x272 y256 w48 h23, &Help
Gui Add, Button, x328 y256 w48 h23, &Cancel
Gui Add, Button, x384 y256 w48 h23 Default, &OK
Gui Add, Edit, vStart x24 y56 w184 h21
Gui Add, Button, x216 y56 w33 h22, &File
Gui Add, CheckBox, valldays x184 y100 w70 h19, Every day

Gui Show, w450 h288, AHK-Tasks-Scheduler v1
Return

; SUBROUTINES


ButtonFile:
FileSelectFile, SelectedFile, 3, , Open a file ;Text Documents (*.txt; *.doc)

GuiControl,1:, Start, %SelectedFile%
return


ButtonOK:
	Gui, submit, nohide

	; Format date and time for wake-up-function
	Stringmid, YYYY,StartDate,1,4	;Year
	Stringmid, MM,StartDate,5,2	;Month
	Stringmid, DD,StartDate,7,2	;Day
	Stringmid, Hour,StartTime,9,2	;hour
	Stringmid, Min,StartTime,11,2 ;minute

	;Add "" to application name
	Start="%Start%"


	; Prepare code for standby-type: 0=no standby 1= hibernate 2= suspend (Radio Buttons in a group are numbered!)
	Standby:=Standby-1

	; Prepare code for wake-up: 0=no wake up 1= wake up
	Resume:=WakeUp-1

	; Prepare code for Every days: alldays=0=no alldays=1=yes
    
	
	; Set Timer by running timer script
	;run, autohotkey.exe %WakeUpMain% %YYYY% %MM% %DD% %Hour% %Min% %Standby% %Start% %Resume%
run, autohotkey.exe %WakeUpMain% %YYYY% %MM% %DD% %Hour% %Min% %Standby% %Start% %Resume% %alldays%
    return

Esc::
ButtonCancel:
GuiClose:
	ExitApp	
    
    

ButtonHelp:
helptext=
(  Join`r`n
1-How to install: PutWapUp.ahk file in the same folder
2-Run AHK-Task-scheduler and choose a file to schedule:
    -schedule single task: job is deleted after run
    -schedule every task: job remains minimized in tray and start each day (Every day check)
    
@adegard, 2018 AHK-Tasks-scheduler-v1.ahk
https://github.com/adegard/AHK-Tasks-scheduler
)

	MsgBox, %helptext%
	return
