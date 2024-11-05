'Cria o objeto para controlar as ações das teclas do teclado
'Objeto de automação da classe Word.Basic
Set WBasic= CreateObject("Word.Basic")
'Aciona o botão de print screen
WBasic.sendkeys"{prtsc}"
'Faz o script dormir por 2 segundos para dar tempo de tirar o print
WScript.Sleep 2000
'Cria o objeto para executar (run) o comando mspaint e abrir o paint
set WshShell = WScript.CreateObject("WScript.Shell")
'Utilizamos o método Run para executar o comando mspaint
WshShell.Run "mspaint"
'Faz o script dormir por 2 segundos para dar tempo de abrir o paint
WScript.Sleep 2000
'Após abrir o paint pelo comando mspaint, o mesmo se encontra inativo/minimizado
'Precisamos adentrar na aplicação, através do comando AppActivate
WBasic.AppActivate "Paint"
'Faz o script dormir por 2 segundos para dar tempo de entrar no paint
WScript.Sleep 2000
'Aciona os botões ctrl + v para colar o print no paint
WBasic.sendkeys"^(v)"
'Faz o script dormir por 2 segundos para dar tempo colar o print
WScript.Sleep 2000
'Aciona os botões ctrl + s para salvar o arquivo do paint
WBasic.sendkeys"^(s)"
'Faz o script dormir por 2 segundos para dar tempo de abrir a tela de salvar
WScript.Sleep 2000
'Digita no teclado meuPrint.jpg
WBasic.sendkeys"meuPrint.jpg"
'Faz o script dormir por 2 segundos para dar tempo de digitar o nome do arquivo a ser salvo
WScript.Sleep 2000
'Aciona o botão de enter para concluir a ação de salvar
WBasic.sendkeys"{ENTER}"
'Faz o script dormir por 2 segundos para dar tempo de concluir a ação de salvar
WScript.Sleep 2000
'Aciona os botões ctrl + F4 para fechar o paint
WBasic.sendkeys"%{F4}"
'Faz o script dormir por 2 segundos para dar tempo de fechar o paint
WScript.Sleep 2000
'Exibe mensagem de sucesso na tela
WScript.Echo "Print salvo com sucesso!"