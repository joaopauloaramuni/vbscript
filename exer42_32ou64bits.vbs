Option Explicit
 
Dim intAddressWidth
Dim colItems, objItem, objWMIService
Dim strMsg
 
Set objWMIService = GetObject( "winmgmts://./root/cimv2" )
Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_Processor")

For Each objItem in colItems
	intAddressWidth = objItem.AddressWidth
	strMsg = "Windows" & vbTab & ": " & intAddressWidth   & "-bit" & vbCrLf _
	       & "Processor" & vbTab & ": " & objItem.DataWidth & "-bit"
Next

WScript.Echo strMsg
