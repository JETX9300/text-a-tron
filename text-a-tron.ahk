;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance force	; force the script to reload when the app is started again

; Starting Variables
History0 := "----"
History1 := "----"
History2 := "----"
History3 := "----"
History4 := "----"
History5 := "----"
History6 := "----"
History7 := "----"
History8 := "----"
History9 := "----"

size := 400

noProjection := true
watchList := false
hold := false
WLlatch := true

; If in testing mode don't show projection
if (noProjection)
	{
		Goto, skipProjection
	}

; Check how many Monitors the system has
SysGet, MonCount, MonitorCount

; If more than 1 monitor, ask which monitor to use as a projector
if (MonCount > 1)
	{
	Gui, New, , Text-a-Tron
	Gui, Add, Text, Section, Display on Which Screen?
	Loop, %MonCount%
		{
		Gui, Add, Text, Section xs, Monitor %A_Index%
		Gui, Add, Checkbox, vMonitor%A_Index% ys,
		}
	Gui, Add, Button, Section xs gStart, Start
	Gui, Show
	return
	}

Start:
Gui, Submit
Gui, Destroy

Loop, %MonCount%
	{
		If (Monitor%A_Index%)
			{
				SysGet, Mon%A_Index%, Monitor, %A_Index%
				ProjMonWidth := Mon%A_Index%Right - Mon%A_Index%Left
				ProjMonHeight := Mon%A_Index%Bottom - Mon%A_Index%Top
				ProjMonX := Mon%A_Index%Left
				ProjMonY := Mon%A_Index%Top
			}
	}


;FileRead, filecontent, %A_ScriptFullPath%	; read some file

;Set up Screen for Projection
Gui,+LastFound		; Sets the window to be the last found window, to use for winset ...
Gui,-border			; Provides a thin-line border around the window. This is not common.
Gui,+AlwaysOnTop	; Makes the window stay on top of all other windows
Gui,+ToolWindow		; Provides a narrower title bar but the window will have no taskbar button.
Gui,-Caption		; Provides a title bar and a thick window border/edge. When removing the caption from a window that will use WinSet TransColor, remove it only after setting the TransColor
Gui,Color,black 	; color of the gui ;black ;000000
Gui,Font, s400 Bold		; Set a large font size


Gui,Add,Text, Center x0 y0 cWhite 0x200 W%ProjMonWidth% H%ProjMonHeight% vTicketNumber,%DisplayNumber% ; create the textfield with its content

if (MonCount > 1)
	{
			Gui,Show, x%ProjMonX% y%ProjMonY% W%ProjMonWidth% H%ProjMonHeight%	; Show the gui in fullscreen on selected Monitor
	}
	else
	{
	Gui,Show, x0 y0 W%A_ScreenWidth% H%A_ScreenHeight%
	}

skipProjection:

; Setup Gui
Gui, TAT: New, , Text-a-Tron
Gui, TAT: Add, Text, Section ,Current Number:
Gui, TAT: Add, Text, ys vCurrentNumber w100,
Gui, TAT: Add, Edit, Section xs vVal,
Gui, TAT: Add, Button, gOK ys Default, OK
Gui, TAT: Add, Text, Section xs, History:
Gui, TAT: Add, Text, xs vHist0 w100, ----
Gui, TAT: Add, Text, xs vHist1 w100, ----
Gui, TAT: Add, Text, xs vHist2 w100, ----
Gui, TAT: Add, Text, xs vHist3 w100, ----
Gui, TAT: Add, Text, xs vHist4 w100, ----
Gui, TAT: Add, Text, xs vHist5 w100, ----
Gui, TAT: Add, Text, xs vHist6 w100, ----
Gui, TAT: Add, Text, xs vHist7 w100, ----
Gui, TAT: Add, Text, xs vHist8 w100, ----
Gui, TAT: Add, Text, xs vHist9 w100, ----
Gui, TAT: Add, Button, Section gWList, Watch List
Gui, TAT: Show

;Main Loop
loop
	{
		If (hold = false)
			{
				GuiControlGet, Current,,Val
				StringTrimRight, Col, Current, Strlen(Current) - 1
				If(Col = "R")
					{
					Color = Red
					}
					Else If(Col = "O")
					{
					Color = FFA500
					}
					Else If(Col = "G")
					{
					Color = Green
					}
					Else If(Col = "Y")
					{
					Color = Yellow
					}
					Else If(Col = "B")
					{
					Color = Blue
					}
					Else
					{
					Color = White
					}

				; Text size control
				if(DisplayNumber != Current)
					{
					If(Strlen(Current) < 5)
						{
						Gui, 1:Font, s400 c%Color% Bold
						GuiControl, 1:Font, TicketNumber
						}
					If(Strlen(Current) >= 5 and Strlen(Current) <= 6)
						{
						Gui, 1:Font, s325 c%Color% Bold
						GuiControl, 1:Font, TicketNumber
						}
					If(Strlen(Current) > 6)
						{
						Gui, 1:Font, s250 c%Color% Bold
						GuiControl, 1:Font, TicketNumber
						}
					DisplayNumber := Current
					StringTrimRight, FirstC, Current, Strlen(Current) - 1
					StringTrimLeft, Trimmed, Current, 1
					If(FirstC = "R" or FirstC = "O" or FirstC = "G" or FirstC = "Y" or FirstC = "B")
						{
						GuiControl,,CurrentNumber,%Trimmed%
						GuiControl,1:,TicketNumber,%Trimmed%
						}
						else
						{
						GuiControl,,CurrentNumber,%Current%
						GuiControl,1:,TicketNumber,%Current%
						}
					}
			}
		sleep 100
	}


return				; end of the autoexecute section



OK:
Gui, TAT: Submit, nohide
GuiControlGet, CurrentTicket,,CurrentNumber


;Check if someone on the Watch list has Won
If (watchList and WLlatch)
	{
		Gui, Watchlist:Default
		Loop % LV_GetCount()
			{
				LV_GetText(checkName, A_Index)
				LV_GetText(checkTicket1, A_Index, 2)
				LV_GetText(checkTicket2, A_Index, 3)

				;Msgbox, Current Number:%CurrentTicket% Current Check:%checkName% %checkTicket1% - %checkTicket2%

				If (Strlen(checkTicket2) = 0)
					{
						If (CurrentTicket = checkTicket1)
							{
								Msgbox, Congratulations %checkName%!

								WLlatch := false
							}
					}
					else
					{
						If (CurrentTicket >= checkTicket1 and CurrentTicket <= checkTicket2)
							{
								Msgbox, Congratulations %checkName%!

								WLlatch := false
							}
					}
			}
	}
Gui, TAT:Default

;update History
If (Strlen(Val) > 0)
	{
		History9 = %History8%
		History8 = %History7%
		History7 = %History6%
		History6 = %History5%
		History5 = %History4%
		History4 = %History3%
		History3 = %History2%
		History2 = %History1%
		History1 = %History0%
		History0 = %Val%

		GuiControl,,Hist0,%History0%
		GuiControl,,Hist1,%History1%
		GuiControl,,Hist2,%History2%
		GuiControl,,Hist3,%History3%
		GuiControl,,Hist4,%History4%
		GuiControl,,Hist5,%History5%
		GuiControl,,Hist6,%History6%
		GuiControl,,Hist7,%History7%
		GuiControl,,Hist8,%History8%
		GuiControl,,Hist9,%History9%

		GuiControl,,Val
		GuiControl, Disable, Val
		hold := true
	}
	else
	{
		GuiControl,,CurrentNumber,%Val%
		GuiControl,1:,TicketNumber,%Val%
		hold := false
		WLlatch := true
		GuiControl, Enable, Val
		GuiControl, Focus, Val
	}
Return

WList:
;Create the Watch List Gui
If (watchList = false)
	{
		Gui, WatchList: New,, Watch List
		Gui, WatchList: Add, Text, Section x10 y10, Name
		Gui, WatchList: Add, Text, x+53, First Ticket
		Gui, WatchList: Add, Text, x+12, Last Ticket
		Gui, WatchList: Add, Edit, x10 y25 w70 vNewName,
		Gui, WatchList: Add, Edit, x+10 w55 vTicket1,
		Gui, WatchList: Add, Edit, x+10 w55 vTicket2,
		Gui, WatchList: Add, Button, x+10 gAddTickets Default, Add
		Gui, WatchList: Add, ListView, Section x10 y70 Grid vWatchList r20, Name|First Ticket|Last Ticket
		Gui, WatchList: Add, Button, xs gDeleteRow, Delete Row
		Gui, WatchList: Add, Button, x+108 gImportCsv, Import CSV
		Gui, WatchList: Show

		WatchList := true
	}
Return

AddTickets:
Gui, WatchList: Submit, nohide

;Check to see if the Information has been filled out
If (Strlen(NewName) = 0)
  {
    Msgbox, Please fill out a name to enter

    Return
  }

If (Strlen(Ticket1) = 0)
  {
    Msgbox, Please add at least 1 ticket number

    Return
  }

;Add new values to list
LV_Add(,NewName,Ticket1,Ticket2)

;Clear Text Input Boxes
GuiControl,,NewName,
GuiControl,,Ticket1,
GuiControl,,Ticket2,

;move cursor back to the first field
GuiControl, Focus, NewName

LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")
LV_ModifyCol(3, "AutoHdr")
LV_ModifyCol(1,"Sort")
Gui, WatchList: Show, AutoSize

Return

DeleteRow:
Gui, Watchlist:Default
toDelete := LV_GetNext(1,"F")
LV_Delete(toDelete)
Gui, TAT:Default
Return

ImportCsv:
InputBox, path, Import CSV, Full path of CSV file to use?,,,,,,,,%A_Desktop%\
If ErrorLevel
	{
		Return
	}
	else
	{
		If (Strlen(path) = 0)
			{
				Msgbox, Please enter a file path
				goto ImportCsv
			}

		Loop, read, %path%
			{
				If (A_Index > 1)
					{
						Loop, parse, A_LoopReadLine, CSV
							{
								If (A_Index = 1)
									{
										NewName := A_LoopField
									}
								If (A_Index = 2)
									{
										Ticket1 := A_LoopField
									}
								If (A_Index = 3)
									{
										Ticket2 := A_LoopField
									}
							}
						LV_Add(,NewName,Ticket1,Ticket2)
					}
					LV_ModifyCol(1, "AutoHdr")
					LV_ModifyCol(2, "AutoHdr")
					LV_ModifyCol(3, "AutoHdr")
					LV_ModifyCol(1,"Sort")
					Gui, WatchList: Show, AutoSize
			}
	}
Return

1GuiEscape:     ; if you press escape in the gui
TATGuiEscape:
WatchListGuiEscape:
1GuiClose:
TATGuiClose:
WatchListGuiClose:
exitapp
Return
