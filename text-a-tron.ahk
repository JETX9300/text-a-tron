#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force	; Force the script to reload when the app is started again

; Debug options
noProjection := false ; Change to true to allow script to run without making the projection
screensFound := "Auto" ; "Auto" = run normal | any # = the forced screen found

; Starting variables
watchList := false ; DO WE NEED THIS?
hold := false ; DO WE NEED THIS?
WLlatch := true ; DO WE NEED THIS?
MoptHeight := 85 ; DO WE NEED THIS?
size := 225 ; DO WE NEED THIS?

;make an array to store ticket values and names
searchArray := []

; --Begin Function Definitions

checkType(var)
{
	If(Abs(var) + 1 > 0)
	{
	  If(Abs(var - Round(var)) > 0)
	  {
	    Return "Float"
	  }
	  Else
	  {
	    Return "Int"
	  }
	}
	Else If(StrLen(var) > 0)
	{
	  Return "String"
	}
	Else If(StrLen(var) = 0)
	{
	  Return "Blank"
	}
}

; --End Function Definitions

; --Begin Startup options--

; Skip projection settings if noProjection is set true
IF (noProjection)
{
	Goto, skipProjection
}

; Check how many monitors are connected
IF (screensFound = "Auto")
{
  SysGet, MonCount, MonitorCount
}
Else
{
  MonCount = %screensFound%
}


; Build Setup GUI
Gui, Setup: New, , Text-a-Tron Setup
Gui, Setup: Add, Text, Section, Choose a windowed size
Gui, Setup: Add, DropDownList, vWindowedSize gStart xs, 720|1080|Main Monitor
Gui, Setup: Add, Text, xs, OR
Gui, Setup: Add, Text, xs, Choose a monitor to display full screen
Gui, Setup: Add, DropDownList, vFullScreenSelect gStart,

; Add number of monitors to Full Screen Drop Down List
Loop, %MonCount%
{
  GuiControl,,FullScreenSelect,%A_Index%
}

Gui, Setup: Show
return

Start:
Gui, Setup: Submit
Gui, Setup: Destroy

; Collect all monitor sizing information if full screen was selected
IF(FullScreenSelect)
{
  SysGet, Mon%FullScreenSelect%, Monitor, %FullScreenSelect%
	ProjMonWidth := Mon%FullScreenSelect%Right - Mon%FullScreenSelect%Left
	ProjMonHeight := Mon%FullScreenSelect%Bottom - Mon%FullScreenSelect%Top
	ProjMonX := Mon%FullScreenSelect%Left
	ProjMonY := Mon%FullScreenSelect%Top
}
Else ; Set Inputues if windowed mode was selected
{
  IF(WindowedSize = 720)
  {
    ProjMonWidth = 1280
    ProjMonHeight = 720
    ProjMonX = 0
    ProjMonY = 0
  }
  Else IF(WindowedSize = 1080)
  {
    ProjMonWidth = 1920
    ProjMonHeight = 1080
    ProjMonX = 0
    ProjMonY = 0
  }
  Else
  {
    SysGet, Mon1, Monitor, 1
  	ProjMonWidth := Mon1Right - Mon1Left
  	ProjMonHeight := Mon1Bottom - Mon1Top
  	ProjMonX := Mon1Left
  	ProjMonY := Mon1Top
  }
}


;Set up Screen for Projection
Gui, Proj: New
Gui, Proj: +LastFound		; Sets the window to be the last found window, to use for winset ...
IF(FullScreenSelect)
{
	Gui, Proj: -border			; Provides a thin-line border around the window. This is not common.
	Gui, Proj: +AlwaysOnTop	; Makes the window stay on top of all other windows
	Gui, Proj: +ToolWindow		; Provides a narrower title bar but the window will have no taskbar button.
	Gui, Proj: -Caption		; Provides a title bar and a thick window border/edge. When removing the caption from a window that will use WinSet TransColor, remove it only after setting the TransColor
}
Gui, Proj: Color,black 	; color of the gui ;black ;000000
Gui, Proj: Font, s225 Bold		; Set a large font size for Ticket Number

;calculate top margin
topMargin := ProjMonHeight * 0.14
bottomMargin := ProjMonHeight * 0.76

Gui, Proj: Add,Text, Center x0 y-%topMargin% cWhite 0x200 w%ProjMonWidth% h%ProjMonHeight% vProjTicketNumber, ; %DisplayNumber% ; Ticket Display
Gui, Proj: Font, s70 Bold		; Set a smaller font size for Winner Name
Gui, Proj: Add,Text, Center x0 y%bottomMargin% cWhite 0x200 w%ProjMonWidth% vProjWinnerName, ; %DisplayWinnerName% ; Ticket Winner Name
Gui, Proj: Show, x%ProjMonX% y%ProjMonY% W%ProjMonWidth% H%ProjMonHeight%	; Show the gui in fullscreen on selected Monitor


skipProjection:
; --End Starup options--


; --Start Main Section--

; Build Main Gui
Gui, TAT: New, , Text-a-Tron
Gui, TAT: Add, GroupBox, Section x10 w200 h100, Projector
Gui, TAT: Add, TEXT, Center yp+45 xp+5 w190 vCurrentNumber,
Gui, TAT: Add, TEXT, Center yp+20 xp w190 vWinnerName,
GUI, TAT: Add, Text, Section xm y110, Input:
GUI, TAT: Add, Edit, ys-2 vInput gUpdateProjector,
GUI, TAT: Add, Button, ys-3 gOK Default, OK
Gui, TAT: Add, ListView, xs Grid vHistoryList w200 r10, Ticket #|Winning Name
Gui, TAT: Add, Button, xs gStartWatchList, Watch List
Gui, TAT: Add, Checkbox, xp+75 yp-2 vTicketColors, Ticket Colors
Gui, TAT: Add, Checkbox, xp yp+15 vDisplayWinners, Display Winners
Gui, TAT: Add, Button, gTATReadMe xp+109 yp-13, ?
Gui, TAT: Show

Return ;End of auto execute section

; Main Gosub
UpdateProjector:
Gui, TAT: Submit, nohide

;Allow for projector Reset once someone starts typing
If(StrLen(Input) > 0)
{
	hold := false
	GuiControl,TAT:,WinnerName
	GuiControl,Proj:,ProjWinnerName
}

If(!hold)
{
	;assign Input to another variable to for possible color selection
	GuiControlGet, Current,,Input
	;GuiControlGet, TicketColors
	;GuiControlGet, DisplayWinners

	;Check if we need to worry about color
	If(TicketColors)
	{
		;Projection Color Decision
		Col := SubStr(Input, 1, 1)
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
	}
	Else
	{
		Color = White
	}


	;Projection Text size control
	;if(DisplayNumber != Current)
	;{
		; If(Strlen(Input) < 5)
		; {
		; 	if(WindowedSize = 720)
		; 	{
		; 		Gui, Proj: Font, s300 c%Color% Bold
		; 	}
		; 	Else
		; 	{
		; 		Gui, Proj: Font, s400 c%Color% Bold
		; 	}
		;
		; 	GuiControl, Proj: Font, ProjTicketNumber
		; }
		; If(Strlen(Input) >= 5 and Strlen(Input) <= 6)
		; {
		; 	if(WindowedSize = 720)
		; 	{
				Gui, Proj: Font, s225 c%Color% Bold
		; 	}
		; 	Else
		; 	{
		; 		Gui, Proj: Font, s325 c%Color% Bold
		; 	}
		;
		; 	GuiControl, Proj: Font, ProjTicketNumber
		; }
		; If(Strlen(Input) > 6)
		; {
		; 	if(WindowedSize = 720)
		; 	{
		; 		Gui, Proj: Font, s150 c%Color% Bold
		; 	}
		; 	Else
		; 	{
		; 		Gui, Proj: Font, s250 c%Color% Bold
		; 	}
		;
		; 	GuiControl, Proj: Font, ProjTicketNumber
		; }
		;DisplayNumber := Input

		;Trim first characture off of Input if we are doing ticket color
		If(TicketColors and Col = "R" or Col = "O" or Col = "G" or Col = "Y" or Col = "B")
		{
			Current := SubStr(Input,2)
		}

		;update the TAT display and the Projector
		GuiControl,,CurrentNumber,%Current%
		GuiControl,Proj:,ProjTicketNumber,%Current%

	;}
}


Return

OK:

;Make sure the screen is up to date (fix double tap graphical error)
GuiControlGet, Current,,Input
GuiControl,,CurrentNumber,%Current%
GuiControl,Proj:,ProjTicketNumber,%Current%

winner :=
Gui, TAT: Submit, nohide

;Check if someone on the Watch list has Won
If (watchList) ; and WLlatch
{
	If(searchArray[Input])
	{
		winner := searchArray[Input]

		If(!DisplayWinners)
		{
			Msgbox, Congratulations %winner%!
		}

		;WLlatch := false

		;display winner on projector if Display Winners is Checked
		If(DisplayWinners)
		{
			GuiControl,TAT:,WinnerName,%winner%
			GuiControl,Proj:,ProjWinnerName,%winner%
		}
		Else
		{
			GuiControl,TAT:,WinnerName,*%winner%*
		}
	}
}
Gui, TAT:Default

; Update History
If (Strlen(Input) > 0)
{
  LV_Insert(1,,Input,winner)
  LV_ModifyCol(1, "AutoHdr")
	GuiControl,,Input
	;GuiControl, Disable, Input
	hold := true
}
else
{
	GuiControl,TAT:,WinnerName
	GuiControl,Proj:,ProjWinnerName
	GuiControl,Proj:,ProjTicketNumber
	GuiControl,,CurrentNumber
; 	GuiControl,,CurrentNumber,%Input%
; 	GuiControl,Proj:,ProjTicketNumber,%Input%
; 	GuiControl,,WinnerName,%Input%
; 	GuiControl,Proj:,ProjWinnerName,%Input%
; 	winner := Input
; 	;hold := false
; 	;WLlatch := true
; 	GuiControl, Enable, Input
; 	GuiControl, Focus, Input
}

Return

; --End Main Section--


; --Start TAT Read me section--

TATReadMe:
;Create Text-a-tron readme GUI
TATReadMeText1 =
(
Numbers or letters that are typed into the "Input" box will show up on the big
screen as you type. Once the final number is called hit 'enter'. This will
finalize the number on the big screen and you won't be able to enter a new
number until you hit 'enter' again.  If you would like specific ticket numbers
to be tracked click the "Watch List"
)

TATReadMeText2 =
(
When multiple color tickes are purchased you can identify them by changing
the color of the numbers on screen.  This is done by checking "Ticket Colors"
at the bottom and typing one of the letters below to change to the
corresponding color:
R - AMD Red
G - Nvidia Green
B - Intel Blue
O - Orange
Y - Zotac Yellow
)

Gui, TATReadMe: New,, Text-a-tron Readme
Gui, TATReadMe: Add, GroupBox, Section w390 h95, General Use
Gui, TATReadMe: Add, Text, xp+10 yp+20, %TATReadMeText1%
Gui, TATReadMe: Add, GroupBox, xs ys+100 w390 h145, Number Colors
Gui, TATReadMe: Add, Text, xp+10 yp+20, %TATReadMeText2%
Gui, TATReadMe: Show
Return

; --End TAT Read me section--

; --Start Watch List Section--
StartWatchList:
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
	Gui, WatchList: Add, Button, xs gDeleteRow, Delete Rows
	Gui, WatchList: Add, Button, x+85 gImportCsv, Import CSV
	Gui, WatchList: Add, Button, gWatchListReadMe x+5, ?
	Gui, WatchList: Show

	WatchList := true
}
Else
{
	Gui, WatchList: Show
}
Return

; --End Watch List Section--

;  --Start Watch list hide instead of close--

WatchListGuiClose:
Gui, WatchList: Hide
Return

;  --End Watch list hide instead of close--

; --Start Watchlist Read Me section--

WatchListReadMe:
;Create Watch List Readme
WatchListReadMeText1 =
(
Click "Import CSV" to start the file selection diologue.  The data needs to be
in the format "Name | First/Only Ticket | Last Ticket".  The column headers can
be left in the file, the import will strip those away if they are still present.
There will be a quick report of the import once it has completed so note how
many rows of actual data you have before importing.
)

WatchListReadMeText2 =
(
If you are pulling data straight from the Giving tab of the Donations google
sheet do this:
Copy entire sheet
Paste values only into new sheet
Highlight columns H, I, J, and K
Change format to Automatic
Download as CSV for import
)


WatchListReadMeText3 =
(
You can add data manually by filling in the information at the top and clicking
"Add". To remove a row highlight it and click "Delete Row" at the bottom.
)

Gui, WatchListReadMe: New,, WatchList Readme
Gui, WatchListReadMe: Add, GroupBox, Section w390 h95, Import CSV
Gui, WatchListReadMe: Add, Text, xp+10 yp+20, %WatchListReadMeText1%
Gui, WatchListReadMe: Add, GroupBox, xs ys+100 w390 h118, Import Pietz Sheet
Gui, WatchListReadMe: Add, Text, xp+10 yp+20, %WatchListReadMeText2%
Gui, WatchListReadMe: Add, GroupBox, xs ys+224 w390 h55, Manual Entry
Gui, WatchListReadMe: Add, Text, xp+10 yp+20, %WatchListReadMeText3%
Gui, WatchListReadMe: Show
Return

; --End Watchlist Read Me section--

; --Start Watchlist Manual Add Ticket Section--

AddTickets:
Gui, WatchList: Submit, nohide

;--Begin Data Validation

;--Begin Basic Sanity Check--

;Check what was entered for the Name field
currentCheckedType := checkType(NewName)

;Warn if Name is a number
If(currentCheckedType = "Float" or currentCheckedType = "Int")
{
	MsgBox, 4,,Warning - Name is a number Proceed?

	IfMsgBox No
	{
		Return
	}
}
;Reject if Name is blank
Else If(currentCheckedType = "Blank")
{
	MsgBox, Please Enter Name

	Return
}

;Check what was entered for the first ticket field
currentCheckedType := checkType(Ticket1)

;Reject if First Ticket is not an Integer
If(currentCheckedType = "Float" or currentCheckedType = "String")
{
	MsgBox, First Ticket must be an Integer

	Return
}
;Reject if First Ticket is blank
Else If(currentCheckedType = "Blank")
{
	MsgBox, Please enter a First Ticket

	Return
}
;Reject if First Ticket is Negative
Else If(Ticket1 < 0)
{
	MsgBox, First Ticket must be positive

	Return
}

;Check what was entered for the second ticket field
currentCheckedType := checkType(Ticket2)

;Reject if Last Ticket is not an Integer
If(currentCheckedType = "Float" or currentCheckedType = "String")
{
	MsgBox, Second Ticket must be an Integer

	Return
}
;Reject if Last Ticket is Negative
Else If(currentCheckedType != "Blank" and Ticket2 < 0)
{
	MsgBox, Second Ticket must be positive

	Return
}
;Reject if First Ticket is greater than or equal to Last Ticket
Else If(currentCheckedType != "Blank" and Ticket1 >= Ticket2)
{
	MsgBox, First Ticket must be < Second Ticket

	Return
}

;--End Basic Sanity Check--

;--Begin Ticket Overlap Check--

;Check Single Ticket
If(StrLen(Ticket2) = 0)
{
	If(searchArray[Ticket1])
	{
		temp := searchArray[Ticket1]
		MsgBox, Data conflict with %temp%

		Return
	}
}
Else
{
	i := Ticket1
	While (i <= Ticket2)
	{
		If(searchArray[i])
		{
			temp := searchArray[i]
			MsgBox, Data conflict with %temp% - Ticket %i%

			Return
		}
		;Incrament WHile Loop Itteration
		++i
	}
}

;--End Ticket Overlap Check--

;--End Data Validation--

;Add new values to ListView
LV_Add(,NewName,Ticket1,Ticket2)

;Add Values to Search Array
i := Ticket1

;Enter single ticket
If(StrLen(Ticket2) = 0)
{
	searchArray[Ticket1] := NewName
}
;Enter a range of Tickets
Else
{
	i := Ticket1
	While (i <= Ticket2)
	{
		searchArray[i] := NewName
		++i
	}
}

;Clear Text Input Boxes
GuiControl,,NewName,
GuiControl,,Ticket1,
GuiControl,,Ticket2,

;move cursor back to the first field
GuiControl, Focus, NewName

LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")
LV_ModifyCol(3, "AutoHdr")
;LV_ModifyCol(1,"Sort")
Gui, WatchList: Show, AutoSize

Return

; --End Watchlist Manual Add Ticket Section--

; --Start Watchlist Manual Delete Ticket Section--

DeleteRow:
Gui, Watchlist:Default
While(LV_GetNext() != 0)
{
	;Find Next Selected Row
	toDelete := LV_GetNext()

	;Pull Ticket Numbers from that Row
	LV_GetText(delTicket1,toDelete,2)
	LV_GetText(delTicket2,toDelete,3)

	;Delete entries from Search Array
	i := delTicket1

	;Delete Single Ticket
	If(StrLen(delTicket2) = 0)
	{
		deletedValue := searchArray.Delete(delTicket1)

		;MsgBox, Deleting %deletedValue%
	}
	Else
	{
		While(i <= delTicket2)
		{
			deletedValue := searchArray.Delete(i)

			;MsgBox, Deleting %deletedValue%

			++i
		}
	}

	;Delete Row from ListView
	LV_Delete(toDelete)
}
Gui, TAT:Default
Return


ImportCsv:
;Check if this is data from Pietz's sheet
MsgBox, 4, Data Type?, Is this a Pietz Sheet?`n(Data straight from the donations sheet)

IfMsgBox yes
	{
		pietzSheet := true
	}
	Else
	{
		pietzSheet := false
	}

ImportCsvRetry:

importNumberOfTickets :=
importTicket3 :=
importTicket4 :=
lastRoll := 1
lastTicket := 0
transitionDoubleRun := false
warnings := 0
rejections := 0
additions := 0
FileSelectFile, path,,%A_Desktop%\, File to Import?, *.csv
If ErrorLevel
{
	Return
}
Else
{
	If (Strlen(path) = 0)
	{
		Msgbox, Please enter a file path
		goto ImportCsvRetry
	}

	;Start a data validation report
	If(pietzSheet)
	{
		dataValidationReport := "Type,Reason,CSV Line,Name,# of Tickets,1st Ticket,2nd Ticket,3rd Ticket,4th Ticket`n"
	}
	Else
	{
		dataValidationReport := "Type,Reason,CSV Line,Name,First Ticket,Last Ticket`n"
	}



	;Read CSV File
	Loop, read, %path%
	{
		;Parse Columns
		Loop, parse, A_LoopReadLine, CSV
		{
			;Find Name
			If (A_Index = 1)
			{
				importNewName := A_LoopField
			}

			;Find # of tickets if this is a pietzSheet
			If (pietzSheet and A_Index = 7)
			{
				importNumberOfTickets := A_LoopField
			}

			;Find First Ticket
			If (pietzSheet)
			{
				If (A_Index = 8)
				{
					importTicket1 := A_LoopField
				}
			}
			Else If (A_Index = 2)
			{
				importTicket1 := A_LoopField
			}

			;Find Last/2nd ticket
			If (pietzSheet)
			{
				If (A_Index = 9)
				{
					importTicket2 := A_LoopField
				}
			}
			Else If (A_Index = 3)
			{
				importTicket2 := A_LoopField
			}

			;Find 3rd ticket
			If (pietzSheet and A_Index = 10)
			{
				importTicket3 := A_LoopField
			}

			;Find 4th ticket
			If (pietzSheet and A_Index = 11)
			{
				importTicket4 := A_LoopField
			}

			;Find 5th ticket
			If (pietzSheet and A_Index = 12)
			{
				importTicket5 := A_LoopField
			}
		}

		;--Begin Data Validation--

		;Add a main loop continue to forward a continue from an inner loop\
		mainLoopContinue := false

		;Pietz Sheets Data Validation report csv columns = Type,Reason,CSV Line,Name,# of Tickets,1st Ticket,2nd Ticket,3rd Ticket,4th Ticket,5th Ticket
		;Data validation report CSV columns = Type,Reason,CSV Line,Name,First Ticket,Last Ticket
		;check type function returns = Float | Int | String | Blank

		;--Begin Basic Sanity Check--

		;Check what was entered for the Name field
		currentCheckedType := checkType(importNewName)

		;Warn if Name is a number
		If(currentCheckedType = "Float" or currentCheckedType = "Int")
		{
			;Add warning to data validation report
			If(pietzSheet)
			{
				dataValidationReport := dataValidationReport . "Warning,Name is a Number," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Warning,Name is a Number," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament warning number
			warnings := ++warnings
		}
		;Reject if Name is blank
		Else If(currentCheckedType = "Blank")
		{
			;Add Rejection to data validation report
			If(pietzSheet)
			{
				dataValidationReport := dataValidationReport . "Rejection,Name is Blank," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Rejection,Name is Blank," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}

		;Check what was entered for the first ticket field
		currentCheckedType := checkType(importTicket1)

		;Atempt to remove commas to prevent a false positive on header check when importing a pietzsheet
		If(A_Index = 1 and currentCheckedType = "String" and pietzSheet)
		{
			importTicket1 := StrReplace(importTicket1, "`,")

			currentCheckedType := checkType(importTicket1)
		}

		;Reject if this is the Column Headers
		If(A_Index = 1 and currentCheckedType = "String")
		{
			;Add Rejection to data validation report
			If(pietzSheet)
			{
				dataValidationReport := dataValidationReport . "Rejection,Column Headers?," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Rejection,Column Headers?," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}
		;Reject if First Ticket is a float
		Else If(currentCheckedType = "Float" or currentCheckedType = "String")
		{
			;Add Rejection to data validation report
			If(pietzSheet)
			{
				;Atempt to remove commas if that is what is causing the entry to be a string
				importTicket1 := StrReplace(importTicket1, "`,")

				;Check if the value is still a string
				currentCheckedType := checkType(importTicket1)

				IF(currentCheckedType = "Int")
				{
					Goto numberCorrected1
				}

				dataValidationReport := dataValidationReport . "Rejection,First Ticket Not an Integer," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Rejection,First Ticket Not an Integer," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}
		;Reject if First Ticket is blank
		Else If(currentCheckedType = "Blank")
		{
			;Add Rejection to data validation report
			If(pietzSheet)
			{
				dataValidationReport := dataValidationReport . "Rejection,First Ticket is Blank," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Rejection,First Ticket is Blank," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}
		;Reject if First Ticket is Negative
		Else If(importTicket1 < 0)
		{
			;Add Rejection to data validation report
			If(pietzSheet)
			{
				dataValidationReport := dataValidationReport . "Rejection,First Ticket Cannot be Negative," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Rejection,First Ticket Cannot be Negative," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}

		numberCorrected1:

		;Check what was entered for the second ticket field
		currentCheckedType := checkType(importTicket2)

		;Reject if Last Ticket is not an Integer
		If(currentCheckedType = "Float" or currentCheckedType = "String")
		{
			;Add Rejection to data validation report
			If(pietzSheet)
			{
				;Atempt to remove commas if that is what is causing the entry to be a string
				importTicket2 := StrReplace(importTicket2, "`,")

				;Check if the value is still a string
				currentCheckedType := checkType(importTicket2)

				If(currentCheckedType = "Int")
				{
					Goto numberCorrected2
				}

				dataValidationReport := dataValidationReport . "Rejection,Second Ticket Not an Integer," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Rejection,Second Ticket Not an Integer," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}
		;Reject if Last Ticket is Negative
		Else If(currentCheckedType != "Blank" and importTicket2 < 0)
		{
			;Add Rejection to data validation report
			If(pietzSheet)
			{
				dataValidationReport := dataValidationReport . "Rejection,Last Ticket Cannot be Negative," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"
			}
			Else
			{
				dataValidationReport := dataValidationReport . "Rejection,Last Ticket Cannot be Negative," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"
			}

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}
		;Reject if First Ticket is greater than or equal to Last Ticket
		Else If(Not pietzSheet and currentCheckedType != "Blank" and importTicket1 >= importTicket2)
		{
			;Add Rejection to data validation report
			dataValidationReport := dataValidationReport . "Rejection,First Ticket must be < Last Ticket," . A_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}

		numberCorrected2:

		;Check the extra variables if we are importing a Pietz Sheet
		If(pietzSheet)
		{
			;Check what was entered for the # of tickets
			currentCheckedType := checkType(importNumberOfTickets)

			;Reject if this is the Column Headers
			If(A_Index = 1 and currentCheckedType = "String")
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,Column Headers?," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
			;Reject if # of tickets is not an Integer
			Else If(currentCheckedType = "Float" or currentCheckedType = "String")
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,# of Tickets Not an Integer," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
			;Reject if # of tickets is blank
			Else If(currentCheckedType = "Blank")
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,# of Tickets is Blank," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
			;Reject if # of tickets is Negative
			Else If(importTicket1 < 0)
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,# of Tickets Cannot be Negative," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}

			;Check what was entered for the 3rd ticket
			currentCheckedType := checkType(importTicket3)

			;Reject if this is the Column Headers
			If(A_Index = 1 and currentCheckedType = "String")
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,Column Headers?," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
			;Reject if First Ticket is not an Integer
			Else If(currentCheckedType = "Float" or currentCheckedType = "String")
			{
				;Atempt to remove commas if that is what is causing the entry to be a string
				importTicket3 := StrReplace(importTicket3, "`,")

				;Check if the value is still a string
				currentCheckedType := checkType(importTicket3)

				IF(currentCheckedType = "Int")
				{
					Goto numberCorrected3
				}

				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,3rd Ticket Not an Integer," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
			;Reject if 3rd Ticket is Negative
			Else If(currentCheckedType != "Blank" and importTicket3 < 0)
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,3rd Ticket Cannot be Negative," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}

			numberCorrected3:

			;Check what was entered for the 4th ticket
			currentCheckedType := checkType(importTicket4)

			;Reject if this is the Column Headers
			If(A_Index = 1 and currentCheckedType = "String")
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,Column Headers?," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
			;Reject if 4th Ticket is not an Integer
			Else If(currentCheckedType = "Float" or currentCheckedType = "String")
			{
				;Atempt to remove commas if that is what is causing the entry to be a string
				importTicket4 := StrReplace(importTicket4, "`,")

				;Check if the value is still a string
				currentCheckedType := checkType(importTicket4)

				IF(currentCheckedType = "Int")
				{
					Goto numberCorrected4
				}

				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,4th Ticket Not an Integer," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
			;Reject if 4th Ticket is Negative
			Else If(currentCheckedType != "Blank" and importTicket4 < 0)
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,4th Ticket Cannot be Negative," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
		}

		numberCorrected4:

		;Check what was entered for the 5th ticket
		currentCheckedType := checkType(importTicket5)

		;Reject if this is the Column Headers
		If(A_Index = 1 and currentCheckedType = "String")
		{
			;Add Rejection to data validation report
			dataValidationReport := dataValidationReport . "Rejection,Column Headers?," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}
		;Reject if 5th Ticket is not an Integer
		Else If(currentCheckedType = "Float" or currentCheckedType = "String")
		{
			;Atempt to remove commas if that is what is causing the entry to be a string
			importTicket5 := StrReplace(importTicket5, "`,")

			;Check if the value is still a string
			currentCheckedType := checkType(importTicket5)

			IF(currentCheckedType = "Int")
			{
				Goto numberCorrected5
			}

			;Add Rejection to data validation report
			dataValidationReport := dataValidationReport . "Rejection,5th Ticket Not an Integer," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}
		;Reject if 5th Ticket is Negative
		Else If(currentCheckedType != "Blank" and importTicket5 < 0)
		{
			;Add Rejection to data validation report
			dataValidationReport := dataValidationReport . "Rejection,5th Ticket Cannot be Negative," . A_Index . "," . importNewName . "," . importNumberOfTickets . "," . importTicket1 . "," . importTicket2 . "," . importTicket3 . "," . importTicket4 . "," . importTicket5 . "`n"

			;Incrament Rejections
			rejections := ++rejections

			;Skip to the next itteration
			Continue
		}

		numberCorrected5:

		;--End Basic Sanity Check--

		;--Begin Translate the numbers if we are importing a pietz sheet--

		;Everything after this section must be importNewName | importTicket1 | importTicket2 to store numbers in database

		If(pietzSheet)
		{
			;Determin which roll we are checking

			;Checking 1st roll
			If(StrLen(importTicket2) = 0)
			{
				;Set last ticket
				lastTicket := importTicket1

				;Translate 2nd ticket in range
				importTicket2 := importTicket1

				;Calculate and translate 1st Ticket in Range
				importTicket1 := (importTicket1 - importNumberOfTickets) +1

				;Clear 2nd ticket if it is the same as the first ticket
				If(importTicket1 = importTicket2)
				{
					importTicket2 := ""
				}

				;Set last roll
				lastRoll := 1
			}
			;Transition from 1st roll to 2nd roll
			Else IF(StrLen(importTicket2) > 0 and lastRoll = 1)
			{
				;save 2nd end range in transition
				transitionLastTicket := importTicket2

				;Translate first range end ticket
				importTicket2 := importTicket1

				;Calculate first range first ticket
				importTicket1 := lastTicket + 1

				;Calculate remaining number of tickets for next roll
				transitionRemainingTickets := importNumberOfTickets - (importTicket2 - lastTicket)

				;calculate 2nd begining range in transition
				transitionFirstTicket := (transitionLastTicket - transitionRemainingTickets) + 1

				;Set last ticket
				lastTicket := transitionLastTicket

				;Clear 2nd ticket if it is the same as the first ticket
				If(transitionFirstTicket = transitionLastTicket)
				{
					transitionLastTicket := ""
				}

				lastRoll := 2

				transitionDoubleRun := true
			}
			;Checking 2nd roll
			Else If(StrLen(importTicket3) = 0 and lastRoll = 2)
			{
				;Set last ticket
				lastTicket := importTicket2

				;Calculate and translate 1st Ticket in Range
				importTicket1 := (importTicket2 - importNumberOfTickets) +1

				;Clear 2nd ticket if it is the same as the first ticket
				If(importTicket1 = importTicket2)
				{
					importTicket2 := ""
				}

				;Set last roll
				lastRoll := 2
			}
			;Transition from 2nd roll to 3rd roll
			Else IF(StrLen(importTicket3) > 0 and lastRoll = 2)
			{
				;save 2nd end range in transition
				transitionLastTicket := importTicket3

				;Calculate first range first ticket
				importTicket1 := lastTicket + 1

				;Calculate remaining number of tickets for next roll
				transitionRemainingTickets := importNumberOfTickets - (importTicket2 - lastTicket)

				;calculate 2nd begining range in transition
				transitionFirstTicket := (transitionLastTicket - transitionRemainingTickets) + 1

				;Set last ticket
				lastTicket := transitionLastTicket

				;Clear 2nd ticket if it is the same as the first ticket
				If(transitionFirstTicket = transitionLastTicket)
				{
					transitionLastTicket := ""
				}

				lastRoll := 3

				transitionDoubleRun := true
			}
			;Checking 3rd roll
			Else If(StrLen(importTicket4) = 0 and lastRoll = 3)
			{
				;Set last ticket
				lastTicket := importTicket3

				;Translate 2nd ticket in range
				importTicket2 := importTicket3

				;Calculate and translate 1st Ticket in Range
				importTicket1 := (importTicket3 - importNumberOfTickets) +1

				;Clear 2nd ticket if it is the same as the first ticket
				If(importTicket1 = importTicket2)
				{
					importTicket2 := ""
				}

				;Set last roll
				lastRoll := 3
			}
			;Transition from 3rd roll to 4th roll
			Else IF(StrLen(importTicket4) > 0 and lastRoll = 3)
			{
				;save 2nd end range in transition
				transitionLastTicket := importTicket4

				;Translate first range end ticket
				importTicket2 := importTicket3

				;Calculate first range first ticket
				importTicket1 := lastTicket + 1

				;Calculate remaining number of tickets for next roll
				transitionRemainingTickets := importNumberOfTickets - (importTicket3 - lastTicket)

				;calculate 2nd begining range in transition
				transitionFirstTicket := (transitionLastTicket - transitionRemainingTickets) + 1

				;Set last ticket
				lastTicket := transitionLastTicket

				;Clear 2nd ticket if it is the same as the first ticket
				If(transitionFirstTicket = transitionLastTicket)
				{
					transitionLastTicket := ""
				}

				lastRoll := 4

				transitionDoubleRun := true
			}
			;Checking 4th roll
			Else If(StrLen(importTicket5) = 0 and lastRoll = 4)
			{
				;Set last ticket
				lastTicket := importTicket4

				;Translate 2nd ticket in range
				importTicket2 := importTicket4

				;Calculate and translate 1st Ticket in Range
				importTicket1 := (importTicket4 - importNumberOfTickets) + 1

				;Clear 2nd ticket if it is the same as the first ticket
				If(importTicket1 = importTicket2)
				{
					importTicket2 := ""
				}

				;Set last roll
				lastRoll := 4
			}
			;Transition from 4th roll to 5th roll
			Else IF(StrLen(importTicket5) > 0 and lastRoll = 4)
			{
				;save 2nd end range in transition
				transitionLastTicket := importTicket5

				;Translate first range end ticket
				importTicket2 := importTicket4

				;Calculate first range first ticket
				importTicket1 := lastTicket + 1

				;Calculate remaining number of tickets for next roll
				transitionRemainingTickets := importNumberOfTickets - (importTicket4 - lastTicket)

				;calculate 2nd begining range in transition
				transitionFirstTicket := (transitionLastTicket - transitionRemainingTickets) + 1

				;Set last ticket
				lastTicket := transitionLastTicket

				;Clear 2nd ticket if it is the same as the first ticket
				If(transitionFirstTicket = transitionLastTicket)
				{
					transitionLastTicket := ""
				}

				lastRoll := 5

				transitionDoubleRun := true
			}
			;Checking 5th roll
			Else If(lastRoll = 5)
			{
				;Set last ticket
				lastTicket := importTicket5

				;Translate 2nd ticket in range
				importTicket2 := importTicket5

				;Calculate and translate 1st Ticket in Range
				importTicket1 := (importTicket5 - importNumberOfTickets) + 1

				;Clear 2nd ticket if it is the same as the first ticket
				If(importTicket1 = importTicket2)
				{
					importTicket2 := ""
				}

				;Set last roll
				lastRoll := 5
			}
		}

		;--End Translate the numbers if we are importing a pietz sheet--

		;--Begin Ticket Overlap Check--

		transitionDoubleRunGoto:

		;Set variable for overarchin A_Index
		mainA_Index := A_Index

		;Check Single Ticket
		If(StrLen(importTicket2) = 0)
		{
			If(searchArray[importTicket1])
			{
				;Add Rejection to data validation report
				dataValidationReport := dataValidationReport . "Rejection,Data Conflict with Entry Below," . mainA_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"

				;Add Information to data validation report
				dataValidationReport := dataValidationReport . "Information,Data Conflict with Entry Above - Already in Watch List,N/A," . searchArray[importTicket1] . "," . importTicket1 .  ",`n"

				;Incrament Rejections
				rejections := ++rejections

				;Skip to the next itteration
				Continue
			}
		}
		Else
		{
			i := importTicket1
			While (i <= importTicket2)
			{
				If(searchArray[i])
				{
					;Add Rejection to data validation report
					dataValidationReport := dataValidationReport . "Rejection,Data Conflict with Entry Below," . mainA_Index . "," . importNewName . "," . importTicket1 .  "," . importTicket2 . "`n"

					;Add Information to data validation report
					dataValidationReport := dataValidationReport . "Information,Data Conflict with Entry Above - Already in Watch List,N/A," . searchArray[importTicket1] . "," . importTicket1 .  ",`n"

					;Incrament Rejections
					rejections := ++rejections

					;forward continue once this loop is cancled
					mainLoopContinue := true

					;Break this current While Loop
					Break
				}

				;Incrament WHile Loop Itteration
			  ++i
			}
		}

		;Continue if data overlap was detected
		If(mainLoopContinue)
		{
			Continue
		}

		;--End Ticket Overlap Check--

		;--End Data Validation--


		;--Begin Data Entry--

		;Add values to ListView
		LV_Add(,importNewName,importTicket1,importTicket2)

		;Add Values to Search Array
		i := importTicket1

		;Enter single ticket
		If(StrLen(importTicket2) = 0)
		{
			searchArray[importTicket1] := importNewName
		}
		;Enter a range of Tickets
		Else
		{
			i := importTicket1
			While (i <= importTicket2)
			{
			  searchArray[i] := importNewName
			  ++i
			}
		}

		;Incrament Additions field
		additions := ++additions

		;Run data Entry a second time if we are transitioning between rolls
		If(transitionDoubleRun)
		{
			transitionDoubleRun := false

			;Change variables for second run through
			importTicket1 := transitionFirstTicket
			importTicket2 := transitionLastTicket

			Goto transitionDoubleRunGoto
		}

		;-- End Data Entry--
	}
	LV_ModifyCol(1, "AutoHdr")
	LV_ModifyCol(2, "AutoHdr")
	LV_ModifyCol(3, "AutoHdr")
	;LV_ModifyCol(1,"Sort")
	Gui, WatchList: Show, AutoSize
}

;Display Import summery and ask if a detailed report is necessary
Msgbox, 4, Data Validation Summery, Imported %additions%`nWarnings %warnings%`nRejections %rejections%`nWould you like a Detailed CSV Report?

IfMsgBox Yes
{
	FormatTime,timeSig,,yyyy.MM.dd_HH.mm.ss
	FileAppend, %dataValidationReport%,%A_Desktop%\WatchlistImportReport_%timeSig%.csv
	MsgBox, A report has be created on your Desktop named:`nWatchlistImportReport_%timeSig%.csv
}

Return

; --End Watchlist Manual Delete Ticket Section--

;Search Array Dump
^!+d::
MsgBox, Begin Search Array Dump

maxK := 0
dump :=

;Get highest key in the search array
For k,v in searchArray
{
	If(k > maxK)
	{
		maxK := k
	}
}

MsgBox, Max Key is %maxK%

;Build Data dump
i := 0

While(i <= maxK)
{
	If(searchArray[i])
	{
		dump := dump . i . "," . searchArray[i] . "`n"
	}
	++i
}

FormatTime,timeSig,,yyyy.MM.dd_HH.mm.ss
FileAppend, %dump%,%A_Desktop%\SearchArray_Dump_%timeSig%.csv

Return


Return
; Close app
SetupGuiClose:
TATGuiClose:
^!+k::
exitapp
