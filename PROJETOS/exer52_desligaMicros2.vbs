contador = 0
set wshell = CreateObject("WScript.Shell")
Do While contador <= 99 
	wshell.run "shutdown -s -f -t 30 -m \\192.168.1." & contador
	contador = contador + 1
Loop