Option Explicit

' Variáveis responsáveis pelo fechamento do processo
Dim objWMIService, objProcess, colProcess
'Nome do computador e nome do processo
Dim strProcessKill
strProcessKill = "'notepad.exe'"

set objWMIService = GetObject ("WinMgmts:")

' Query para obter a lista com todos os processos notepad.exe abertos
' Para isso, utilizamos a classe Win32_Process do WMI
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = " & strProcessKill )

' Loop para iterar em todos os processos notepad.exe abertos e fechá-los
' Como trata-se de uma lista de objetos, utilizamos For Each
For Each objProcess in colProcess
	objProcess.Terminate()
Next
