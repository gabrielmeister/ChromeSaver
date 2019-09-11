;ChromeSaver
;(c)2019 Gabe Meister http://gabemeisterlaw.com/
;Requires AutoIt Windows Scripting Language to use or compile

#include <MsgBoxConstants.au3>
#include <APIDlgConstants.au3>
#include <Memory.au3>
#include <WinAPIDlg.au3>
#include <WinAPIMisc.au3>

$nomore = 0;set stop trigger
While $nomore = 0 
    $outPath = "C:\Users\gabri\OneDrive\Desktop\READ ME";select output folder
    Opt("WintitleMatchMode", 2)
    WinActivate("Google Chrome");activates Chrome window with top tab
    Sleep(400)
    $RawTabTitle = WinGetTitle("Google Chrome");grabs title of top tab (may need to be cleaned to save it)
    $RegExNonStandard="(?i)([^a-z0-9-_])"
    $TabTitle=StringRegExpReplace($RawTabTitle,$RegExNonStandard,"_")
    Send("^+p");select system print dialog (not Chrome's)
    If WinWaitActive("Print", "", 5) = 0 Then Exit;dumps out if something wrong (i know...)
    ControlFocus("Print", "", "ComboBox1");select printer list within dialog
    Send("Microsoft Print to PDF");this is the printer I want
    Send("{ENTER}");start printing
    If WinWait("Save Print Output As", "", 5) = 0 Then Exit;somethin' really bad happened
    Sleep(200)
    ControlFocus("Save Print Output As", "", "Edit1");focuses on Save As edit box
    Sleep(200)
    Send($TabTitle & ".pdf");write the filename and path to Save As edit box
    Send("{ENTER}");saves into outPath
    Sleep(400)
    $wtext = WinGetText("Save Print Output As")
    If StringInStr($wtext, "already exists") Then
    Sleep(1000)
    Send("{Y");ENTER}");yes overwrite
    EndIf
    Sleep(800)
    Beep(450, 300)
    WinActivate("Google Chrome");back to Chrome
    Send("^w");close already-printed tab, which exposes the next tab
    If WinActive("Google Chrome")=0 Then
        $nomore=1
    EndIf;see if there are any more Chrome windows or tabs active; if not, then end While loop
    Sleep(300)
WEnd

	
