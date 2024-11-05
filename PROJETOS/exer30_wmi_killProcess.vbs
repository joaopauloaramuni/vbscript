' Exemplo WMI para matar um processo do windows
' Neste exemplo, vamos fechar todas as calculadoras que estiverem abertas

Option Explicit

' Variáveis responsáveis pelo fechamento do processo
Dim objWMIService, objProcess, colProcess

'Nome do computador e nome do processo
Dim strComputer, strProcessKill
strComputer = "."
strProcessKill = "'calc.exe'"

' Contador para sabermos quantas calculadoras foram fechadas
Dim contCalc
contCalc = 0

' Conexão WMI com o Root CIM
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

' Query para obter a lista com todos os processos calc.exe abertos
' Para isso, utilizamos a classe Win32_Process do WMI
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = " & strProcessKill )

' Loop para iterar em todos os processos calc.exe abertos e fechá-los
' Como trata-se de uma lista de objetos, utilizamos For Each
For Each objProcess in colProcess
	contCalc = contCalc + 1
	objProcess.Terminate()
Next

if contCalc > 0 then

	WSCript.Echo "Matamos o processo " & strProcessKill & " no computador: " & strComputer & VbCr & _ 
	"Total de processos encerrados: " & contCalc
else
	
	WScript.Echo "Nenhum processo "  & strProcessKill & " aberto no momento."
	
End if	

' Sair - Opcional
WScript.Quit
