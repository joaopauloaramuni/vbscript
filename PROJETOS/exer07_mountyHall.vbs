' Porta dos Desesperados

' Inicializa gerador randômico de números
Randomize

dim portaEscolhida 
dim portaSemPremio 
dim portaComPremio 

'Caso a porta escolhida seja também a porta com prêmio, sobrará uma porta
dim portaQueSobrou

' O usuário saberá que uma das portas não possui prêmio, e poderá escolher entre a porta premiada e a porta que já havia escolhido.
' É claro, sem saber em qual das duas portas está o prêmio.
dim trocarPorta

' Sorteia a porta com o prêmio
portaComPremio = Int( 3 * Rnd + 1 )

' Escolhe uma porta (33,3% de chance de acertar de cara)
portaEscolhida = InputBox("Escolha uma porta: [ 1 ] [ 2 ] [ 3 ]")

' Sorteia um número de 1 a 3 até que a condição seja atendida (ser o único número não escolhido pelo usuário nem pelo sistema)
Do
	portaSemPremio = Int( 3 * Rnd + 1 )

' Para realizarmos os testes, é necessário converter as portas para inteiro, pois este é o tipo de dado gerado pela Randomize
Loop While Int(portaSemPremio) = Int(portaComPremio) or Int(portaSemPremio) = Int(portaEscolhida)

' Ok. Agora temos a porta com o prêmio, a porta que o usuário escolheu, e a porta não escolhida que está sem prêmio

' Estatísticamente falando, sua chance de acertar trocando a porta é de 66,6%, enquanto a chance de ganhar o prêmio permanecendo
' na mesma porta escolhida ao início, é de apenas 33,3%. A partir do momento que uma porta é revelada como sem prêmio, uma nova
' informação é inserida no algoritmo, dobrando suas chances de vencer.

if Int(portaEscolhida) = Int(portaComPremio) Then
	Do
		portaQueSobrou = Int( 3 * Rnd + 1 )
	Loop While Int(portaQueSobrou) = Int(portaComPremio) or Int(portaQueSobrou) = Int(portaSemPremio)

	'Neste caso, o usuário está com a porta com prêmio, se trocar de porta irá perder o prêmio.
	trocarPorta = InputBox("A porta: [" & portaSemPremio & "] não tem prêmio. Deseja trocar a porta [" & portaEscolhida & "] pela porta: [" & portaQueSobrou & "] ? S / N") 
	
	If trocarPorta = "S" Then
		portaEscolhida = portaQueSobrou

	End If

else
	
	'Neste caso, o usuário não está com a porta com prêmio, se trocar de porta irá ganhar o prêmio.
	trocarPorta = InputBox("A porta: [" & portaSemPremio & "] não tem prêmio. Deseja trocar a porta [" & portaEscolhida & "] pela porta: [" & portaComPremio & "] ? S / N") 
	
	If trocarPorta = "S" Then
		portaEscolhida = portaComPremio
	End If
	
End If

'Dois jeitos de vencer, permanecendo com a porta ou trocando de porta
If Int(portaEscolhida) = Int(portaComPremio) and trocarPorta = "N" Then
	WScript.Echo ("VOCÊ GANHOU !!! O prêmio estava na porta: [" & portaComPremio & "]. Deu sorte pois a chance de ganhar era de apenas 33,3% neste caso.")
Elseif Int(portaEscolhida)  = Int(portaComPremio) and trocarPorta = "S" Then
	WScript.Echo ("VOCÊ GANHOU !!! O prêmio estava na porta: [" & portaComPremio & "]. Boa escolha ao trocar! A chance de ganhar era de 66,6% neste caso.")
'Dois jeitos de perder, permanecendo com a porta ou trocando de porta
Elseif Int(portaEscolhida) <> Int(portaComPremio) and trocarPorta = "N" Then
    WScript.Echo ("VOCÊ PERDEU !!! O prêmio estava na porta: [" & portaComPremio & "]. Deveria ter trocado de porta pois a chance era de 66% neste caso.")
Elseif Int(portaEscolhida) <> Int(portaComPremio) and trocarPorta = "S" Then
	WScript.Echo ("VOCÊ PERDEU !!! O prêmio estava na porta: [" & portaComPremio & "]. Se deu mal, não deveria ter trocado de porta, apesar de que as chances eram maiores, 66,6% de chance de ganhar. Você caiu nos outros 33,3%.")
Else
	WScript.Echo ("Erro! Dados foram inseridos incorretamente pelo usuário.")
End If