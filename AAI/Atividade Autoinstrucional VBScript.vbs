
'Autor: Bernardo Mozelli de Medeiros
'Curso: Redes de Computadores
'Discipina: AAI Desenvolvimento de Scripts I
'Professor: João Paulo Aramuni

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

On error resume next

Dim strComputer, strCreate, strNextLine, strMsg1, Str, Str2, Str3, Str4, Str5, Str6, Str7, strFolder
Dim result, result2, result3
Dim FileNameZip, FolderToZip, resultMessage, FolderToBeCreated, deletefolder, sCurPath, strFinish, strFinish2

Const ForReading = 1

strComputer =  InputBox ("Digite o nome do arquivo txt: ", _
						"Autoinstrucional - AAI Desenvolvimento de Scripts I", "Digite o nome do arquivo aqui ")

set objArq = CreateObject("Scripting.FileSystemObject")

If objArq.FileExists (strComputer) then
   result = MsgBox ("O arquivo de texto existe. Deseja prosseguir com a execucao do script?", vbyesNo, "Autoinstrucional - AAI Desenvolvimento de Scripts I")
  
select Case result
Case vbYes
    MsgBox("Prosseguindo com a execucao..."), vbInformation
	WScript.sleep 700

Case vbNo
    MsgBox("Abortando a execucao..."), vbExclamation
	
	WScript.sleep 700
	
	Wscript.Quit
	
	End Select

else

MsgBox "O arquivo de texto nao existe!", vbCritical

Wscript.Quit

end if

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objTextFile = objFSO.OpenTextFile(strComputer,ForReading)

'Le linha por linha do arquivo
Do Until objTextFile.AtEndOfStream
	strNextLine = objTextFile.Readline
	
	If Not objTextFile.ReadLine = false then

		MsgBox "Arquivo aberto com sucesso!", vbInformation

		WScript.sleep 700
		
		else 
		
		MsgBox "Erro ao abrir o arquivo", vbCritical

		Wscript.quit
		
	    End If
		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Utiliza o WMI para verificar o usuario logado, o nome do computador e o dominio. Depois, salva em um txt
			
set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & _
	strNextLine & "\root\cimv2")
			
	loop
        
	Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
        
	For Each objComputer in colSettings

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Cria uma pasta 

Set ObjFolder = CreateObject ("Scripting.FileSystemObject")

FolderToBeCreated = "AAI Desenvolvimento de Scripts I"

				If ObjFolder.FolderExists(FolderToBeCreated) Then
				
				wscript.sleep 1000

				'Avisa se a pasta ja existir
				Str = "Pasta "
				Str = Str & FolderToBeCreated
				Str = Str & " ja existe   "
				MsgBox(str), vbExclamation
				
				wscript.sleep 700

else

'Cria a pasta
 ObjFolder.CreateFolder(FolderToBeCreated)

 strPasta = "AAI Desenvolvimento de Scripts I"
 
'Mostra a mensagem de sucesso apos a pasta ser criada
Str = "A pasta "
Str = Str & strPasta
Str = Str & " foi criada com sucesso"
MsgBox(str), vbinformation

wscript.sleep 700

End If

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	

'Cria o arquivo para salvar as informações de usuario dos computadores da lista
Set objArq = CreateObject("Scripting.FileSystemObject")
Set strCreate = objArq.CreateTextFile("AAI Desenvolvimento de Scripts I\Usuarios_Logados.txt", True)

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Verifica se o arquivo foi criado. Caso tenha sido, exibe a mensagem informando que o foi criado. Caso contrario, exibe mensagem de erro e encerra a execução do script
			
				
				set objArq2 = CreateObject("Scripting.FileSystemObject")
				
				If not objArq2.FileExists ("AAI Desenvolvimento de Scripts I\Usuarios_Logados.txt") then
			
				MsgBox "Erro ao tentar criar o arquivo!", vbCritical
			
				Wscript.Quit
				
				else
	
				MsgBox "Arquivo Usuarios_Logados.txt criado com sucesso!", vbInformation
			
				wscript.sleep 700
				
				strMsg1 = "Nome do Computador: " 
				strMsg2 = "Usuario Logado : " 
				strMsg3 = "Dominio: " 
				strMsg4 = "-------------------------------------------------------------------------------------"
			
				strCreate.WriteLine(strMsg1) & objComputer.Name & vbcrl
				strCreate.WriteLine(strMsg2) & objComputer.username & vbcrl
				strCreate.WriteLine(strMsg3) & objComputer.Domain & vbcrl
				strCreate.WriteLine(strMsg4) 
			
			
			end if 
        
		Next
		
		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
'Zipa a pasta que contem o arquivo Usuarios_Logados.txt

Set FSO = WScript.CreateObject ("Scripting.FileSystemObject")

'Local e nome da pasta que será compactada
FolderToZip = "AAI Desenvolvimento de Scripts I"

'Local e nome do arquivo zip que será criado
FileNameZip = "AAI_VBScript.zip"

'Criando arquivo zip vazio

Set Otf = Fso.OpenTextFile(FileNameZip,2,True)
Otf.Write ""
Otf.Close

'Fecha o objeto
Set Otf = Nothing

Set Fso = Nothing

'Criando objeto shell
Set Shell = CreateObject("Shell.Application")

'Copiando todos os arquivos da pasta para o arquivo zip
Shell.NameSpace(FileNameZip).CopyHere Shell.NameSpace(FolderToZip).Items, &H0&

wscript.sleep 700

'Mensagem de sucesso

Str2 = "Arquivo "
Str2 = Str2 & FileNameZip
Str2 = Str2 & " criado com sucesso"
MsgBox(Str2), vbInformation

'Fecha o objeto
Set Shell = Nothing

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Envia o email com o arquivo zipado em anexo

strName = "Autor: Bernardo Mozelli de Medeiros"

Const schema   = "http://schemas.microsoft.com/cdo/configuration/"
Const cdoBasic = 1
Const cdoSendUsingPort = 2
Dim oMsg, oConf
 
' Propriedades do email
Set oMsg      = CreateObject("CDO.Message")
oMsg.From     = "teste.vbscript@gmail.com" 	'"Nome do remetente <from@gmail.com>. Esse email é valido. Mude conforme sua necessidade"
oMsg.To       = "bernardommedeiros13@gmail.com"       '"Nome do destino <to@gmail.com>. Mude conforme sua necessidade"
oMsg.Subject  = "AAI Desenvolvimento de Scripts I"
oMsg.TextBody = "Arquivo AAI_VBScript.zip enviado com sucesso :)!!!" & vbcrlf & vbCrlf & strName 
oMsg.AddAttachment "C:\Users\Bernardo\Desktop\AAI_VBScript.zip" 
 
'Configuração e autenticação do seu servidor de SMTP Gmail
Set oConf = oMsg.Configuration
 
'Endereço do servidor de SMTP
oConf.Fields(schema & "smtpserver")       = "smtp.gmail.com"
 
'Número da porta
oConf.Fields(schema & "smtpserverport")   = 465
 
oConf.Fields(schema & "sendusing")        = cdoSendUsingPort
 
'Tipo de autenticacao
oConf.Fields(schema & "smtpauthenticate") = cdoBasic
 
'Uso da Encriptação SSL
oConf.Fields(schema & "smtpusessl")       = True
 
'Envia username
oConf.Fields(schema & "sendusername")     = "teste.vbscript@gmail.com"
 
'Envia password
oConf.Fields(schema & "sendpassword")     = "128579539421" 'Senha do email que enviara a mensagem. Não se esqueça de mudar caso o email remetente também mude
 
oConf.Fields.Update()
 
' Envia mensagem
oMsg.Send()

MsgBox "Enviando o email...", vbInformation

wscript.sleep 700
 
' Retorna o status da mensagem
							
							If Err = true then
							resultMessage = "Erro ao enviar o email " & Err.Number & ": " & Err.Description
							Err.Clear()
							
							WScript.sleep 500
							
							Else
							
							resultMessage = "Email enviado com sucesso !!!"
							
							End If
							
							Wscript.sleep 700
	
							MsgBox(resultMessage), vbInformation
												
							Str3 = "Enviado por: " & vbcrlf & Str4 & oMsg.From & vbcrlf & vbCrlf
							
							Str4 = "Para: " & Str4 & oMsg.To
							
							MsgBox "" & Str3 & Str4,vbInformation

							'Fecha o objeto
							Set oMsg = Nothing
							
							Set oConf = Nothing
							
							WScript.sleep 1100
					
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Deleta o arquivo zipado e a pasta criados anteriormente pelo script, caso queira
									
		result2 = MsgBox ("Deseja excluir o arquivo AAI_VBScript.zip?" ,vbyesNo, "Autoinstrucional - AAI Desenvolvimento de Scripts I")
  
		WScript.sleep 700
		
		select Case result2
		Case vbYes
		MsgBox("Excluindo...")
				
		WScript.sleep 700
		
		
Set objDelete = CreateObject("Scripting.FileSystemObject")
			
									objDelete.DeleteFile (FileNameZip)
			
									If Not objDelete.FileExists(FileNameZip) then
			
									Wscript.sleep 700
			
									Str5 = "Arquivo "
									Str5 = Str5 & FileNameZip
									Str5 = Str5 & " deletado com sucesso"
									MsgBox(Str5), vbInformation
							
									WScript.sleep 700
						
									'Set objExc = Nothing
									
						
									else
						
									MsgBox "Erro ao excluir o arquivo " & FileNameZip, vbCritical
			
			
									Wscript.sleep 700
				
									'Fecha o objeto
									objDelete = Nothing 
									objDelete.close
			
									end if
														
		Case vbNo
					
		End Select		
														
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Criei essa variavel para armazenar somente o nome da pasta. Portanto, quando a mesma for concatenada junto ao MsgBox, so ira aparecer o nome da pasta, sem o caminho da mesma.

strFolder = """AAI Desenvolvimento de Scripts I"""
	
result3 = MsgBox ("Deseja excluir a pasta AAI Desenvolvimento de Scripts I?" ,vbyesNo, "Autoinstrucional - AAI Desenvolvimento de Scripts I")

WScript.sleep 700
		
		select Case result3
		
		'Caso a desisão do usuario seja a de não ecluir os arquivos e a pasta, então o script é finalizado
		Case vbNo
		Wscript.Quit
		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		
		Case vbYes
		
		MsgBox("Excluindo...")
		
		wscript.sleep 500
		
		MsgBox "Devido a problemas de permissao no windows, teremos de utilizar um script .bat para apagar a pasta " & strFolder, vbExclamation
		
		wscript.sleep 700
		
		MsgBox "Por favor aguarde!", vbInformation
		
		wscript.sleep 700

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Rem Inicia o Notepad

set WshShell = Wscript.CreateObject("WScript.Shell")

WshShell.run "notepad"
WScript.Sleep 200

WshShell.AppActivate "Notepad"

WScript.Sleep 1000

WshShell.SendKeys "@echo off" & vbCrlf

WshShell.SendKeys ":aa" & vbcrlf

WshShell.SendKeys "MSG * Pasta AAI Desenvolvimento de Scripts I apagada com sucesso !" & vbCrlf  

WshShell.SendKeys "rd /s /q " & strFolder & vbcrlf

WshShell.SendKeys "del remove_pasta.bat" 

Set WBasic = CreateObject("Word.Basic")

'Aciona os botões ctrl + s para salvar o arquivo
WBasic.sendkeys"^(s)"

'Faz o script dormir por 2 segundos para dar tempo de abrir a tela de salvar
WScript.Sleep 2000

'Digita no teclado remove_pasta.bat
WBasic.sendkeys"remove_pasta.bat"
Wscript.Sleep 2000

'Aciona o botão de enter para concluir a ação de salvar
WBasic.sendkeys"{ENTER}"

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
strFinish = "."

'Finaliza o notepad
Set objNetwork = CreateObject("Wscript.Network")

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strFinish & "\root\cimv2")

''' Processo que será verificado '''''''
Set colProcesses = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'notepad.exe'")

''' elimina o processo definido '''
For each Processo in ColProcesses
  Processo.Terminate()
Next			

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
							
wscript.sleep 2000

'Chama o script bat criado para excluir a pasta

Set objWsh = CreateObject ("WScript.Shell")
 
objWsh.Run "remove_pasta.bat"
								
End Select


