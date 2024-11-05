'Obtém o nome do Computador:
set WshNetwork=Wscript.CreateObject("Wscript.Network")
strComputerName=WshNetwork.ComputerName
Wscript.Echo strComputerName