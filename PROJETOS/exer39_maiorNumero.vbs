'Declaração das Variáveis
Dim n1
Dim n2
Dim n3

'Entrada de Dados
n1 = InputBox("Insira um numero:")
n2 = InputBox("Insira um numero:")
n3 = InputBox("Insira um numero:")

'Conversão para inteiro
n1 = CInt(n1) 
n2 = CInt(n2) 
n3 = CInt(n3) 

'Comparações
if ( n1 > n2 and n1 > n3 ) then
	WScript.Echo "Resultado: " & n1 & " é o maior numero."
    elseif ( n2 > n1 and n2 > n3 )  then
	WScript.Echo "Resultado: " & n2 & " é o maior numero."
    elseif ( n3 > n1 and n3 > n2 )  then
	WScript.Echo "Resultado: " & n3 & " é o maior numero."
    elseif ( n1 = n2 and n1 = n3 and n2 = n3 ) then
	WScript.Echo "Todos os três números são iguais."   
    else
	WScript.Echo "Não é possível verificar qual é o maior."  
End if
