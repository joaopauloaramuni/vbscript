Option Explicit
On Error Resume Next

Dim HD, hdc

set HD = WSCript.CreateObject("Scripting.FileSystemObject")

set hdc = HD.GetDrive (HD.GetDriveName("C:"))

Exibe_Msg()

Function Exibe_Msg()

'Considerando um HD de 1 TB = 1048576 MB
MsgBox ("O espaço livre no HD é = " & FormatNumber(hdc.FreeSpace / 1048576, 0) & " MB")

End Function