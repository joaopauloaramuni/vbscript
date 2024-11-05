'Rem Inicia o Notepad
set WshShell= Wscript.CreateObject("WScript.Shell")
WshShell.run "notepad"
WScript.Sleep 200
WshShell.AppActivate "Notepad"
WScript.Sleep 200
WshShell.SendKeys "Top !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! "
Wscript.Sleep 2000