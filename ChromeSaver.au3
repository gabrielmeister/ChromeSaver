;for Windows and requires AutoIt scripting language (until you compile it)
;replace ************** with actual save path
;(c)2019 Gabriel Meister

#include <MsgBoxConstants.au3>
#include <APIDlgConstants.au3>
#include <Memory.au3>
#include <WinAPIDlg.au3>
#include <WinAPIMisc.au3>

$nomore = 0
While $nomore = 0 
    $outPath = "C:**************";select output path and folder
    Opt("WintitleMatchMode", 2)
    WinActivate("Google Chrome")
	Sleep(400)
	$TabTitle = WinGetTitle("Google Chrome")
    Send("^+p");select traditional (i.e., non-Chrome-specific) print dialog
    If WinWaitActive("Print", "", 5) = 0 Then Exit
    ControlFocus("Print", "", "ComboBox1");select the printer list combo
    Send("Microsoft Print to PDF");this is the printer I want
    Send("{ENTER}");start printing
    ConsoleWrite("Waiting for Save As" & @CRLF)
    If WinWait("Save Print Output As", "", 5) = 0 Then Exit;because something bad happened
    Sleep(200)
    ControlFocus("Save Print Output As", "", "Edit1")
    Sleep(200)
	Send($TabTitle & ".pdf");write the filename and path
    Send("{ENTER}");start it saving
    Sleep(400)
    $wtext = WinGetText("Save Print Output As")
    If StringInStr($wtext, "already exists") Then
        Sleep(1000)
        Send("{Y");ENTER}");yes overwrite
    EndIf
    Sleep(2000)
    Beep(450, 300)
	WinActivate("Google Chrome");back to Chrome
    Send("^w");close that tab because it's printed and we'll do the next
	If WinActive("Google Chrome")=0 Then
        $nomore=1
	EndIf 
	Sleep(400)
WEnd

	
