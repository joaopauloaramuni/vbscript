Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives

caminhoArquivoTxt = "U:\meusDrives.txt"
Set arquivoTxt = objFSO.CreateTextFile(caminhoArquivoTxt, True)

For Each objDrive in colDrives
    arquivoTxt.write "Drive letter: " & objDrive.DriveLetter & VbCrLf
Next

arquivoTxt.close