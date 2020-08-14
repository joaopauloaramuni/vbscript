'Script para listar os processos em execução e salvar essa lista em um arquivo txt.

Dim objWMIService, colProcess, objProcess, stringProcess, caminhoArquivoTxt

'Caminho para salvar o arquivo
caminhoArquivoTxt = "C:\Users\Public\processosAbertos.txt"   'Edite essa linha para mudar o caminho do arquivo

'Objeto WMI
Set objWMIService = GetObject("WinMgmts:")

'Obtendo os processos em execução - Classe Win32_Process
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process")
  stringProcess = stringProcess & "Informações dos Processos:" & vbCrLf & vbCrLf
  'Loop para iterar na lista de objetos
  For Each objProcess in colProcess 
    stringProcess = stringProcess & "Processo: " & objProcess.Caption & vbCrLf
    stringProcess = stringProcess & "Caminho do executável: " & objProcess.ExecutablePath & vbCrLf
    stringProcess = stringProcess & "ID do processo: " & objProcess.ProcessID & vbCrLf & "____________________________________________________" & vbCrLf & vbCrLf
	'Inserimos um traço ao final para separar um processo do outro no arquivo .txt
  Next
  
'Limpando os objetos - Opcional
Set colProcess = Nothing
Set objWMIService = Nothing

'Gravando no arquivo .txt
Dim FSO, arquivoTxt
Set FSO = CreateObject("Scripting.FileSystemObject")
Set arquivoTxt = FSO.CreateTextFile(caminhoArquivoTxt, True)
arquivoTxt.Write stringProcess
arquivoTxt.Close

'Mensagem de sucesso
MsgBox "Feito."