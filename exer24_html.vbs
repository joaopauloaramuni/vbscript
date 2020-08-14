' Lista o conteúdo de uma página web

url="http://www.microsoft.com"
Set objHTTP = CreateObject("MSXML2.XMLHTTP")

Call objHTTP.Open("GET", url, FALSE)
objHTTP.Send

WScript.Echo(objHTTP.ResponseText)