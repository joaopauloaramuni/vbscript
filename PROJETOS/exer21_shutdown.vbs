Option Explicit
Dim result
result = MsgBox ("Shutdown?", vbYesNo, "Yes/No Exm")
Select Case result
    Case vbYes
        MsgBox("shuting down ...")
        Dim objShell
        Set objShell = WScript.CreateObject("WScript.Shell")
        objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
    Case vbNo
        MsgBox("Ok")
End Select