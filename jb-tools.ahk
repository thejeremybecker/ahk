#^o::  ; Windows+Control+o hotkey.
IfWinNotExist Inbox - email_address - Outlook
	
    return  ; Outlook isn't open to the right section, so do nothing.
WinActivate  ; Activate the window found by the above command.
Send ^n  ; Create new/blank e-mail via Control+N.
WinWaitActive Untitled Message
Send {Tab} ; Move cursor to body
Send {Tab}
Send {Tab}
;Send {Tab}Dear Sir or Madam,{Enter 2}We have recently discovered a minor defect ...  ; etc.
return  ; This line serves to finish the hotkey.
#SingleInstance force

; -------------------------------------

#^e::  ; Windows+Control+e hotkey.

	Run Excel.exe ;
	WinWaitActive Excel
	Send ^n  ; Create new/blank Excel sheet via Control+N.
    return  ; Excel isn't open to the right section, so run excel.

#SingleInstance force

; -------------------------------------

MoveWindow(width, height)
{
WinMove, A, , , , width, height
ToolTip, %width%x%height%
Sleep, 500
ToolTip,
Return
}
<#1::MoveWindow(1200, 860) ; that's not "standard", just my whole screen
<#2::MoveWindow(1100, 720)
<#3::MoveWindow(800, 600)
<#4::MoveWindow(700, 500)

; -------------------------------------

#v::  ; Windows+v hotkey - to paste values in Excel
	
	WinActivate Excel.exe
	Send !e	
	send s
	Send v
	send {enter}

Return
