Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection

' Ejetar CD-ROM
colCDROMs.Item(0).Eject

WScript.Echo("Ejetado!")