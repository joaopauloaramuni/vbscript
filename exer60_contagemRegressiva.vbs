'Chamada das funções
regressiva("10")

Function regressiva(repetir)
	For i=repetir to 1 step -1 
		'Som
		beep()
		'Luz
		crazyLights()		
		'Fala
		talk(i)
	Next
End Function

Function beep()
	Set oShell = CreateObject("Wscript.Shell")
	oShell.Run "%comspec% /c echo " & Chr(7), 0, False
End Function

Function crazyLights()
	Set oShell = CreateObject("Wscript.Shell")
	oShell.sendkeys "{CAPSLOCK}"
	oShell.sendkeys "{NUMLOCK}"
	oShell.sendkeys "{SCROLLLOCK}"
End Function
	
Function talk(cont)
	Set objVoice = CreateObject("SAPI.SpVoice")
	objVoice.Rate = -5
	objVoice.Speak cont
End Function