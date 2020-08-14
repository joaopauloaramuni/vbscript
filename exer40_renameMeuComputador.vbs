Const MY_COMPUTER = &H11&

Set objNetwork = CreateObject("Wscript.Network")
objComputerName = objNetwork.ComputerName
'objComputerName = "Meu Computador"
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_COMPUTER) 
Set objFolderItem = objFolder.Self
objFolderItem.Name = objComputerName