Dim objWMIService, objProcess, colProcess
const ForReading = 1

set objWMIService = GetObject ("WinMgmts:")
set fso = CreateObject("Scripting.FileSystemObject")
set leia = fso.openTextFile("U:\processos.txt",ForReading)

'Leitura do arquivo processos.txt
Do until leia.AtEndOfStream
processo = leia.readline

	' Query para obter a lista com todos os processos do tipo 'processo' abertos
	' Para isso, utilizamos a classe Win32_Process do WMI
	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & "'" & processo & "'" )

	' Loop para iterar em todos os processos do tipo 'processo' abertos e fechá-los
	' Como trata-se de uma lista de objetos, utilizamos For Each
	For Each objProcess in colProcess
		objProcess.Terminate()
	Next

loop





