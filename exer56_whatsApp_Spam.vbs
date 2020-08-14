' Este script tem a finalidade de enviar mensagens em massa
' para algum contato do seu whatsApp 

' Para realizar o ataque usaremos o whatsApp pela web.
' Para começar a usar o WhatsApp Web:
' Vá até web.whatsapp.com no seu computador.
' Abra o WhatsApp no seu aparelho celular. 
' No Android: vá em Conversas > Menu > WhatsApp Web.
' Escaneie o código da tela do seu computador com o seu telefone.

' InputBox's
Contact = InputBox("Qual contato voce deseja enviar mensagens?", "WhatsApp - Contato")
Message = InputBox("Qual a mensagem que deve ser enviada?","WhatsApp - Mensagem")
T = InputBox("Quantas vezes voce deseja enviar a mensagem?","WhatsApp - Repetir")
If MsgBox("Dados preenchidos corretamente.", 1024 + vbSystemModal, "WhatsApp DDos") = vbOk Then

' Ir para o whatsApp web
Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("https://web.whatsapp.com/", 1)

' Tempo para carregar o whatsApp
If MsgBox("O WhatsApp esta carregado?" & vbNewLine & vbNewLine & "Pressione Nao para cancelar.", vbYesNo + vbQuestion + vbSystemModal, "WhatsApp - Carregado?") = vbYes Then

' Ir para a barra de busca do whatsApp
WScript.Sleep 50
WshShell.SendKeys "{TAB}"

' Ir para o chat de contatos
WScript.Sleep 50
WshShell.SendKeys Contact
WScript.Sleep 50
WshShell.SendKeys "{ENTER}"

' Loop para as mensagens
For i = 0 to T
WScript.Sleep 5
WshShell.SendKeys Message
WScript.Sleep 5
WshShell.SendKeys "{ENTER}"
Next

' Final do Script
WScript.Sleep 3000
MsgBox "O ataque ao contato " + Contact + " esta feito.", 1024 + vbSystemModal, "Ataque realizado."

' Script Cancelado
Else
MsgBox "O processo foi cancelado.", vbSystemModal, "Ataque Cancelado"
End If
Else
End If