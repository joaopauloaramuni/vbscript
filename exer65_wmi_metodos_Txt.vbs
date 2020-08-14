'Classe FSO para criar e escrever no arquivo txt
Set objFSO = CreateObject("Scripting.FileSystemObject")
caminhoArquivoTxt = "U:\metodos_e_propriedades.txt"
'Criar arquivo txt
Set arquivoTxt = objFSO.CreateTextFile(caminhoArquivoTxt, True)

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

'WScript.Echo " Metodos da Classe: " & strClass

arquivoTxt.write "Classe: " & strClass & VbCrLf

arquivoTxt.write VbCrLf & "Métodos:" & VbCrLf & VbCrLf

'Métodos
For Each objClassMethod in objClass.Methods_
    arquivoTxt.write objClassMethod.Name & VbCrLf
Next

arquivoTxt.write VbCrLf & "Propriedades:" & VbCrLf & VbCrLf

'Propriedades
For Each objClassProperty in objClass.Properties_
    arquivoTxt.write objClassProperty.Name & VbCrLf
Next

arquivoTxt.close