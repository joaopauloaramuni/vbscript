'Para testar: 
' 1) Crie uma pasta chamada "pasta" em C:\ 
' 2) Coloque algumas pastas dentro dela e arquivos de texto .txt
' 3) Rode o script e veja que o arquivo backup.zip foi criado em C:\
' 4) Abra o arquivo backup.zip e veja suas pastas e arquivos .txt compactados

'Local e nome da pasta que será compactada
FolderToZip = "C:\pasta"

'Local e nome do arquivo zip que será criado
FileNameZip = "C:\backup.zip"

'Criando arquivo zip vazio
Set Fso = CreateObject("Scripting.FileSystemObject")
Set Otf = Fso.OpenTextFile(FileNameZip,2,True)
Otf.Write ""
Otf.Close
Set Otf = Nothing
Set Fso = Nothing

'Criando objeto shell
Set Shell = CreateObject("Shell.Application")

'Copiando todos os arquivos da pasta para o arquivo zip
Shell.NameSpace(FileNameZip).CopyHere Shell.NameSpace(FolderToZip).Items, &H0&

'Aguardando o fim da compressão
Do Until Shell.NameSpace(FileNameZip).Items.Count = Shell.NameSpace(FolderToZip).Items.Count
   WScript.Sleep 100
Loop

'Mensagem de sucesso
Str = "Zipado! Número de pastas compactadas: "
Str = Str & Shell.NameSpace(FolderToZip).Items.Count
MsgBox(Str)

Set Shell = Nothing