; SETUP STUFF
#NoEnv                     ;Recommended for performance and compatibility with future AutoHotkey releases.
;;SendMode Input           ;I discovered this causes MouseMove to jump as if Speed was 0. (was Recommended for new scripts due to its superior speed and reliability.)
#SingleInstance force      ;Skips the message, "An older instance of this script is already running. Replace it with this instance?"
WinGet, SavedWinId, ID, A  ;Save our current active window
MouseGetPos, xpos, ypos    ;Save initial position of mouse (note: no %% because it's writing output to xpos, ypos)
SetKeyDelay, 60            ;Any number you want (milliseconds)
CoordMode,Mouse,Screen     ;Initial state is Relative
CoordMode,Pixel,Screen     ;Initial state is Relative. Frustration awaits if you set Mouse to Screen and then use GetPixelColor because you forgot this line. There are separate ones for: Mouse, Pixel, ToolTip, Menu, Caret
MouseMove, 0, 0, 0         ;Prevents the status bar from showing a mouse-hover link instead of "Done". (We need to move the mouse out of the way _before_ we go to a webpage.)

;------------------User SetUp-------------------------------------------------
InputBox, plantCode, Plant, Which Plant to Pull Orders From? `r Enter Corresponding Number `r 10: City/La Crosse `r 18: Latrobe `r 28: BCB

;------------------INSTANIATE INTERNET EXPLORER SESSION-----------------------
;URL
	url := "http://vbsapps.pabstbrewingco.com/OrderEntry/EditOrderHeaders.aspx"
; Go to the website in internet explorer
	wb := ComObjCreate("InternetExplorer.Application")
	wb.Visible := true
	wb.Navigate(url)
	WinMaximize, A ;maximize the browser
	; wait until the page loads
	Sleep, 5000
	; below code does not work on every computer for some reason, but is most efficient
	; While	wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy || A_Index < 50
		;Sleep 2

;--------------------PROCESS TO GET ORDERS POPULATED-----------------------
; Checks "Ship Date Range from" box, skip if production month is desired
	idShipDate := "ctl00_PageBodyContentPlaceHolder_rbShipDate"
	wb.document.getElementById(idShipDate).click()
	Sleep, 10

;------------------------------------------------------------------------------

; Selects Plant: 10 for City, 18 for Latrobe, 28 for BCB
	wb.document.all.ctl00_PageBodyContentPlaceHolder_ddlPlant.value := plantCode
	Sleep, 10
; Selects Entered Status
	wb.document.getElementsByClassName("AppDropDown")[6].focus()
	Send {Down}
; Clicks "Get Orders"
	wb.document.getElementsByClassName("AppButton")[0].focus()
	Send {Enter}
;Wait for page to reload
	Sleep, 5000

;--------------------GET DATA INTO EXCEL--------------------------
; copy the page
	Sleep, 8000
	Send ^a
	Sleep, 10
	Send ^c
	Sleep, 10
	
; open excel
	XLBook := ComObjCreate("Excel.Application")
	XLBook.Visible := True
	XLBook.Workbooks.Add
	WinMaximize, A
	Sleep, 30
	
;Paste Page
	Send ^v
	Sleep, 30
	
; Navigate to A107 and copy first order #
	XLBook.Range("A107").Select
	Sleep, 20
	Send ^c
	Sleep, 25

; Switch back to IE
	Send #{Tab}
	Sleep, 60
	Send {Right}
	Sleep, 40
	Send {Enter}
	Sleep, 250

;----------------Print first PDF separately------------------
; Tab to first "View" link
	Send ^f
	Sleep, 10
	Send Edit Plant 
	Sleep, 5
	Send {Esc}
	Sleep, 5
	Send {Tab 5}
	Sleep, 5
	Send {Enter}
	Sleep, 10
	; print first link
	printPage()


;------------------------LOOP THROUGH ALL ORDERS---------------------------------
Loop
{
	;Switch to Excel and copy next order number
		Send #{Tab}
		Sleep, 60
		Send {Right}
		Sleep, 40
		Send {Enter}
		Sleep, 250
		
		;down to next #
		Send {Down}
		Sleep, 5
		clipboard := ""
		Send ^c
		ClipWait
		if(clipboard = "© Pabst Brewing Company 2017, All Rights Reserved`r`n")
		{
			Send !{F4}
			WinWait, ahk_class NUIDialog
			Send {Right}
			Sleep, 5
			Send {Enter}
			MsgBox, Complete
			ExitApp
		}
	; Switch back to IE
		Send #{Tab}
		Sleep, 60
		Send {Right}
		Sleep, 40
		Send {Enter}
		Sleep, 250
	; Tab to "View" link
		Send {Tab 7}
		Sleep, 15
		Send {Enter}
		Sleep, 50
		printPage()
		Sleep, 50
}
---------------------------------------------------------------------------------
;Pause Hotkey Windows + p
#p::Pause

;---------------------PRESS Windows Key + z TO EXIT AT ANY TIME------------------
#z::
	MsgBox, Program exited by user
	ExitApp
return
;--------------------------------------------------------------------------------
	
;--------------------------PRINT THE CURRENT PAGE--------------------------------
printPage()
{
; Call print command
	WinWait, ahk_class Internet Explorer_TridentDlgFrame
	Sleep, 5
	Send ^p
	WinWait, ahk_class #32770, Print 
; Press Enter to confirm print to PDF
	Sleep, 20
	Send {Enter}
	WinWait, Save As
; Write in Order Number as filename
	Send {Tab 5}
	Sleep, 10
	Send ^v
	Sleep, 50
; Press Enter to accept filename for PDF
	Send {Enter}
	Sleep, 100
; Check if File already exists, defaults to not saving it
	IfWinActive, Confirm Save As
	{
		Send {Enter}
		WinWait, Save As				
		Send !{F4}
		WinWaitNotActive, Save As
	}					
; Activate order window
	WinActivate, ahk_class Internet Explorer_TridentDlgFrame
	Sleep, 50
;Press alt + F4 to close order window
	Send !{F4}
	Sleep, 10
	return ;explicit return
}
---------------------------------------------------------------------------------