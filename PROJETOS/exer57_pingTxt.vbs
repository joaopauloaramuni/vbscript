'Script VBS – Pingar IP's lidos de um arquivo .txt
'Variável que representa um ip
dim ip
'Constante que será passada como parâmetro para o método openTextFile da classe FSO
const ForReading = 1

'fso - obj da classe FileSystemObject
set fso = CreateObject("Scripting.FileSystemObject")
'fileTxt - obj que representa o arquivo texto .txt 
'O arquivo texto será retornado através do método openTextFile da classe FSO
set fileTxt = fso.openTextFile("U:\ips.txt",ForReading)
'objShell - obj da classe Shell para acessarmos o método run
Set objShell = WScript.CreateObject("WScript.Shell")

'Loop para ler o arquivo txt até o final
Do until fileTxt.AtEndOfStream	
	'Leitura da linha do txt que possui cada ip
	ip = fileTxt.readLine
	'Execução do comando ping
	objShell.run "ping " & ip
	'Aguardando comando ping ser executado por completo
	Wscript.Sleep 4000
loop
