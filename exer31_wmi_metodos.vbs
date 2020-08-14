'Script para retornar todos os métodos de uma classe wmi
strComputer = "."
strNameSpace = "root\cimv2"

'Descomente e comente as linhas para ver os métodos de cada classe
strClass = "Win32_Service"
'strClass = "Win32_DiskDrive"
'strClass = "Win32_LogicalDisk"
'strClass = "Win32_Process"
'strClass = "Win32_OperatingSystem"

Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\" & strNameSpace & ":" & strClass)

WScript.Echo " Metodos da Classe: " & strClass

For Each objClassMethod in objClass.Methods_
    WScript.Echo objClassMethod.Name
Next

'Propriedades
'For Each objClassProperty in objClass.Properties_
    'WScript.Echo objClassProperty.Name
'Next