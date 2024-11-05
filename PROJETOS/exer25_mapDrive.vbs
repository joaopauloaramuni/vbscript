'Objeto do sistema de arquivos para testar o mapeamento:
Set FSODrive= CreateObject("Scripting.FileSystemObject")

'Testamos para saber se o caminho já não foi mapeado
If not FSODrive.DriveExists("P:") Then
	'Se não foi mapeado, mapeamos...
	'Objeto de mapeamento:
	Set objNetwork = WScript.CreateObject("WScript.Network") 
	'Neste caso, mapiei uma unidade de rede P: para acessar a pasta "relatorios"
	'\\nome_do_servidor\pasta\subpasta...
	objNetwork.MapNetworkDrive "P:", "\\JP-PC\Users\JP\Desktop\relatorios"
	
	'Aqui mudamos o nome da unidade para Relatorios_do_meu_chefe
	Set objShell = CreateObject("Shell.Application")
	objShell.NameSpace("P:").Self.Name = "Relatorios_do_meu_chefe"
	
	'Mensagem de sucesso
	Wscript.Echo "Entre em P: para acessar a pasta de relatorios."

	else
	Wscript.Echo "Unidade ja mapeada!!!"
	
End If