' O ADSI (Active Directory Services Interfaces) é um conjunto de bibliotecas que podem ser utilizadas
' para manipular informações contidas em bases Active Directory. Com o VBScript, podemos utilizar o ADSI
' para manipular as informações do AD diretamente, de forma totalmente personalizada.

' Entre as possibilidades com a utilização das bibliotecas ADSI estão, por exemplo, a verificação de informações sobre
' os usuários cadastrados, a alteração de propriedades e a remoção de entradas não mais necessárias.

' O exemplo de script a seguir armazena em um arquivo todos os usuários cadastrados no Active Directory

' Criar arquivo para escrita
' Arquivo que será utilizado para armazenar as informações obtidas do Active Directory
set objFSO = createObject ("Scripting.FileSystemObject")
const arquivoEscrita = ".\usuarios.txt"
set arqUsuarios = objFSO.createTextFile(arquivoEscrita, True)

' Contar os usuários encontrados
contar = 0

' Criar o objeto que será usado para acessar os registros do Active Directory
' A função getObject("WinNT://exemplo") considera todo o domínio exemplo.local
set objADSI = getObject("WinNT://exemplo")

' O método filter() faz com que sejam considerados apenas os registros de usuário do AD
' (Outras possibilidades são registros de grupos e registros de computadores)
objADSI.filter = Array("User")

' Bloco de repetição que irá percorrer toda a estrutura do AD e escrever no arquivo usuarios.txt
' A propriedade "fullname" corresponde ao nome completo do usuário
' Já a propriedade "name" corresponde ao login de rede desse usuário
for each objUser in objADSI
	arqUsuarios.write objUser.fullname & "-"
	arqUsuarios.write objUser.name & ";" & vbCrlf & vbCrlf
	contar = contar + 1 
next

' Fechando o objeto
arqUsuarios.close

WScript.Echo "Número de usuários encontrados: " & contar

' O arquivo gerado após a execução desse script teria uma forma semelhante à:

' João Francisco de Assis-joaofrancisco;
' Fernando Fernandino-fernando;
' Maria Emilia Boneca-mariaemilia;
' Hugo Quico Chavez-hugo;

' Principais propriedades referentes a usuários cadastros no AD:
' (Há outros atributos que podem estar presentes, uma vez que o Active Directory pode ser 
' facilmente estendido para abrigar novas funcionalidades )

' accountDisabled - true se aconta está desabilitada, false caso contrário
' description - Descrição do usuário
' fullname - Nome completo do usuário
' homeDirectory - Diretório home do usuário
' isAccountLocked - true se a conta está bloqueada, false caso contrário
' loginHours - Horários que o usuário pode se logar
' loginScript - Script de login
' minPasswordLength - Tamanho mínimo da senha de login
' maxPasswordAge - Tempo em que senha de login é válida
' maxPasswordLength - Tamanho máximo de senha de login
' name - Login dos usuários
' passwordHistoryLength - Número de senhas diferentes que devem ser utilizadas antes de ocorrer uma repetição
' profile - Verificar perfil do usuário
' sAMAccountName - Login dos usuários (pré Windows 2000)
' userWorkstations - Estação em que o usuário pode se logar
