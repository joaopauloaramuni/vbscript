'Objeto do sistema de arquivos para testar o mapeamento:
Set FSODrive= CreateObject("Scripting.FileSystemObject")

'Testamos para saber se o caminho está mapeado
If FSODrive.DriveExists("P:") Then
	Set objNetwork = WScript.CreateObject("WScript.Network") 
	'Removemos o mapeamento
	objNetwork.RemoveNetworkDrive "P:"

	
else
	' Como sabemos, o WScript é um objeto intrínseco, ou seja,
	' podemos chamar métodos como Echo e CreateObject de forma direta,
	' sem a necessidade de criarmos um novo objeto para ele
	WScript.Echo "Driver nao mapeado !!!"
	
End If
