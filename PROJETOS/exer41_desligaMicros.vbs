'Script VBS – Desliga micros remotamente

dim computer

set fso = CreateObject("Scripting.FileSystemObject")
const ForReading = 1
set leia = fso.opentextFile("C:\maquinas.txt",ForReading)
Do until leia.AtEndOfStream
computer=leia.readline

set wshell = CreateObject("Wscript.Shell")
wshell.run "shutdown -s -f -t 30 -m \\" & computer

WScript.Echo "shutdown -s -f -t 30 -m \\" & computer
loop

'Obs: Você deve criar um arquivos txt com o ip das máquinas 
'Ex:c:maquinas.txt e infomar este caminho no script