'Documentação:
'https://msdn.microsoft.com/en-us/library/aa394399(v=vs.85).aspx

'Host utilizado: ( . = Próprio host)
strComputer = "."

'Conexão com o WMI
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\Root\CIMv2")

'Classe Win32_ScheduledJob
Set objNewJob = objWMIService.Get("Win32_ScheduledJob")

'Criação da tarefa (Executar Notepad.exe)
'Formato da data: YYYYMMDDHHMMSS.MMMMMM(+-)OOO
errJobCreated = objNewJob.Create("Notepad.exe", "********012500.000000-420", True , 4, , True, JobId) 

'Verificação de Erro
If errJobCreated <> 0 Then
	Wscript.Echo "Erro na criacao da tarefa! Codigo do Erro: " & errJobCreated
Else
	Wscript.Echo "Tarefa criada!"
End If