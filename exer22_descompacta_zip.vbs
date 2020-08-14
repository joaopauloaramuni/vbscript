'Arquivo zipado a ser descompactado
ZipFile="C:\testevbs\teste.zip"
'Pasta onde será extraído o arquivo
ExtractTo="C:\testevbs\top"

'Se o local de extração ainda não foi criado:
Set fso = CreateObject("Scripting.FileSystemObject")
If NOT fso.FolderExists(ExtractTo) Then
   fso.CreateFolder(ExtractTo)
End If

'Extrai o conteúdo do arquivo zip:
set objShell = CreateObject("Shell.Application")
set FilesInZip=objShell.NameSpace(ZipFile).items
objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)
Set fso = Nothing
Set objShell = Nothing