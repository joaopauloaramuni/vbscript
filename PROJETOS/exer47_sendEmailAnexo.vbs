EmailSubject = "Email de teste"
EmailBody = "Este é o corpo do email de teste." & vbCRLF & _
        "Utilizamos o objeto CDO.Message para a autenticação SMTP, através da porta 465."

Const EmailFrom = "joaopauloaramuni@gmail.com"
Const EmailFromName = "My Very Own Name"
Const EmailTo = "joaopauloaramuni@gmail.com"
Const SMTPServer = "smtp.gmail.com"
Const SMTPLogon = "joaopauloaramuni"
Const SMTPPassword = "SENHA"
Const SMTPSSL = True
Const SMTPPort = 465

Const cdoSendUsingPickup = 1    'Send message using local SMTP service pickup directory.
Const cdoSendUsingPort = 2  'Send the message using SMTP over TCP/IP networking.

Const cdoAnonymous = 0  ' No authentication
Const cdoBasic = 1  ' BASIC clear text authentication
Const cdoNTLM = 2   ' NTLM, Microsoft proprietary authentication

' First, create the message

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = EmailSubject
objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
objMessage.To = EmailTo
objMessage.TextBody = EmailBody
objMessage.AddAttachment "C:\pastadeteste\teste.zip" 

' Second, configure the server

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update
'Now send the message!
On Error Resume Next
objMessage.Send

If Err.Number <> 0 Then
    MsgBox Err.Description,16,"Erro ao enviar o email!"
Else 
    MsgBox "Email enviado com sucesso!",64,"Information"
End If