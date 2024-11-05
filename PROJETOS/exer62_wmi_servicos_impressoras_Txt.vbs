'Cria o objFSO que é uma instância pra classe FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Conexão com o WMI através do serviço WinMgmts
set WMI = GetObject ("WinMgmts:")

'Duas variáveis primitivas, servicos e impressoras
Dim servicos
Dim impressoras

'Cria a lista de objetos (objs) que está representando
'todos os servicos rodando na máquina
'Isto é feito através da classe Win32_Service (Classe do WMI)
'Como ela foi acessada? A classe foi instanciada 
'através do método InstancesOf
set objs = WMI.InstancesOf("Win32_Service")

'Loop para objetos (For each)
'Para cada obj contido na lista objs
'(Para cada servico contido na lista de servicos)
'Será recuperado o nome do serviço através da PROPRIEDADE name
'(Informação recuperada: Nome do serviço)
'Os nomes dos serviços serão concatenados na
'variável primitiva 'servicos'
for each obj in objs
	servicos = servicos & obj.name & vbcrlf
next

'Para cada obj contido na lista objs
'(Para cada impressora contido na lista de impressoras)
'Será recuperado o nome da impressora através 
' da PROPRIEDADE name
'(Informação recuperada: Nome da impressora)
'Os nomes das impressoras serão concatenados na
'variável primitiva 'impressoras'
set objs = WMI.InstancesOf("Win32_Printer")
for each obj in objs
	impressoras = impressoras & obj.name & vbcrlf
next

'Neste momento, temos as variáveis 'servicos' e 'impressoras'
'preenchidas com todos os servicos e impressoras mapeados

'Exibem o resultado sem barra de rolagem
WScript.Echo servicos
WScript.Echo impressoras

'Variavel para guardar o caminho do txt onde iremos gravar os servicos
caminhoArquivoTxtServicos = "U:\servicos.txt"
'Criação do obj arquivoTxtServicos que representa o arquivo de texto servicos.txt
'Parametro do método CreateTextFile da Classe FSO: Caminho do arquivo, True (Permite edição)
Set arquivoTxtServicos = objFSO.CreateTextFile(caminhoArquivoTxtServicos, True)

'Variavel para guardar o caminho do txt onde iremos gravar as impressoras
caminhoArquivoTxtImpressoras = "U:\impressoras.txt"
'Criação do obj arquivoTxtImpressoras que representa o arquivo de texto impressoras.txt
'Parametro do método CreateTextFile da Classe FSO: Caminho do arquivo, True (Permite edição)
Set arquivoTxtImpressoras = objFSO.CreateTextFile(caminhoArquivoTxtImpressoras, True)

'Escreve no txt 'servicos.txt' todos os servicos listados
arquivoTxtServicos.write servicos
'Escreve no txt 'impressoras.txt' todos as impressoras listadas
arquivoTxtImpressoras.write impressoras

'Fecha os objetos para evitar que fiquem abertos
arquivoTxtServicos.close
arquivoTxtImpressoras.close