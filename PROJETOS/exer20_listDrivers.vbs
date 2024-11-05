Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives

For Each objDrive in colDrives
    Wscript.Echo "Drive letter: " & objDrive.DriveLetter
Next