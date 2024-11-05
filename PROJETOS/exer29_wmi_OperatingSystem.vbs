' Exemplo WMI para realizar atividades de logoff, shutdown, reboot, etc.
' Neste exemplo, vamos fazer o logoff no sistema

Option Explicit

' Variáveis responsáveis pela chamada dos métodos na classe Win32_OperatingSystem do WMI
Dim objWMIService, oSystem, oSystems

'Nome do computador
Dim strComputer
strComputer = "."

' Conexão WMI com o Root CIM
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

' Query para obter o S.O primário onde estamos logados 
' Para isso, utilizamos a classe Win32_OperatingSystem do WMI
Set oSystems = objWMIService.ExecQuery("select * from Win32_OperatingSystem where Primary=true")

' Loop para iterar em todos os objetos da lista.
' Neste caso teremos apenas um objeto, que é nosso S.O principal/primário onde estamos logados (Primary=true)
' Por padrão, trata-se de uma lista de objetos, por isso utilizamos For Each
For Each oSystem in oSystems
   'LOGOFF   = 0
   'SHUTDOWN = 1
   'REBOOT   = 2
   'FORCE    = 4
   'POWEROFF = 8
   oSystem.Win32Shutdown 0
Next

' Sair - Opcional
WScript.Quit