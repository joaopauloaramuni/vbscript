Set objShell = WScript.CreateObject ("WScript.Shell")

Set objProEnv = objShell.Environment ("Process")

WScript.Echo "Diretório do Windows : " & objProEnv("windir")

WScript.Echo "Caminho do Sistema :" & objProEnv("path")

'Este conjunto de variáveis não retornará nada para o Windows 9x/ME OS.

Set objSysEnv = objShell.Environment("System")

WScript.Echo ("Sistema Operacional : " & objSysEnv("OS"))

WScript.Echo ("Diretório Temp : " & objSysEnv("TEMP"))

WScript.Echo ("Extensões no Path : " & objSysEnv("PATHEXT"))

If objSysEnv ("NUMBER_OF_PROCESSORS") = 1 Then
	WScript.Echo "Seu sistema possui 1 Processador"
Else
	WScript.Echo "Seu sistema possui " & objSysEnv ("NUMBER_OF_PROCESSORS") & " Processadores."
End If	
	