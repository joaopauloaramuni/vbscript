'Inicia calculadora e realiza o cálculo 1 + 2 = 3
set WshShell= Wscript.CreateObject("WScript.Shell")
WsHShell.run "calc"
WScript.Sleep 300
WshShell.AppActivate "Calculator"
WScript.Sleep 300
WshShell.SendKeys "1{+}"
WScript.Sleep 300
WshShell.SendKeys "2"
WScript.Sleep 300
WshShell.SendKeys "~"
WScript.Sleep 300