Option Explicit
Dim objWMIService, objItem, colItems, strComputer, intDrive
strComputer = "."
intDrive = 0

' Conexão WMI com o Root CIM
' O WMI fornece suporte integrado ao modelo CIM (modelo de informação comum),
' o modelo de dados que descreve os objetos existentes em um ambiente de gerenciamento.
' A maioria das classes WMI para gerenciamento estão no namespace root \ cimv2.
' É através destas classes que podemos capturar informações sobre os objetos existentes
Set objWMIService = GetObject("winmgmts:\\" _
& strComputer & "\root\cimv2")

' Query para obter todas as unidades de disco existentes
' A classe WMI Win32_DiskDrive representa uma unidade de disco físico
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive")

' Loop para iterar em todas as unidades de disco existentes
' Utilizamos For Each pois se trata de uma lista de objetos
For Each objItem in colItems
intDrive = intDrive + 1
Wscript.Echo "DiskDrive " & intDrive & vbCr & _ 
"Caption: " & objItem.Caption & VbCr & _ 
"Description: " & objItem.Description & VbCr & _ 
"Manufacturer: " & objItem.Manufacturer & VbCr & _ 
"Model: " & objItem.Model & VbCr & _ 
"Name: " & objItem.Name & VbCr & _ 
"Partitions: " & objItem.Partitions & VbCr & _ 
"Size: " & objItem.Size & VbCr & _ 
"Status: " & objItem.Status & VbCr & _ 
"SystemName: " & objItem.SystemName & VbCr & _ 
"TotalCylinders: " & objItem.TotalCylinders & VbCr & _ 
"TotalHeads: " & objItem.TotalHeads & VbCr & _ 
"TotalSectors: " & objItem.TotalSectors & VbCr & _ 
"TotalTracks: " & objItem.TotalTracks & VbCr & _ 
"TracksPerCylinder: " & objItem.TracksPerCylinder 
Next

'Sair - Opcional
WSCript.Quit
