Option Explicit
Dim contador

contador = 1

WScript.Echo "Vamos contar até 10 !!!"

'Loop irá acontecer até que a condição seja falsa
Do While contador <= 10
	WScript.Echo (contador)
	
	contador = contador +1
	
Loop

'Loop irá acontecer até que a condição seja falsa
'Do
	'WScript.Echo (contador)
	
	'contador = contador +1
	
'Loop While contador <= 10

'Loop irá acontecer até que a condição seja verdadeira
'Do Until contador = 10
	'WScript.Echo (contador)
	
	'contador = contador +1
'Loop
