'O exemplo abaixo cria o diretório testePasta no drive C: e muda a localização para este diretório.

Dim objShell

Set objShell = WScript.CreateObject ("WScript.shell")

objShell.run "cmd /K CD C:\ & mkdir \testePasta & cd testePasta"