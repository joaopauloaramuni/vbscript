' Abrir arquivo de texto com o VBScript e exibir seu conteúdo na tela

' Documentação da Microsoft sobre o FileSystemObject:
' https://msdn.microsoft.com/en-us/library/6tkce7xa(v=vs.84).aspx
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Método OpenTextFile
' iomode - Opção 1: Você irá apenas ler o conteúdo do arquivo, não podendo escrever no mesmo.
Set objFile = objFSO.OpenTextFile("C:\teste.txt", 1)
conteudoArquivo = objFile.ReadAll
Wscript.Echo conteudoArquivo
objFile.Close