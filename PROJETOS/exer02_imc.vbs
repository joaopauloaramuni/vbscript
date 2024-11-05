dim peso
dim altura

peso = InputBox("Digite seu Peso:")
altura = InputBox("Digite seu altura:")
IMC = peso/((altura)*(altura))
MsgBox("Seu IMC é " & IMC)

if ((IMC) < 18.5) Then
	MsgBox "Peso abaixo do ideal!"
elseif (IMC) >= 18.5 and (IMC) <= 24.9 Then
	MsgBox "Peso Ideal!"
elseif (IMC) >= 25 and (IMC) <= 29.9 Then
	MsgBox "Acima do peso ideal!!"
elseif (IMC) >= 30 and (IMC) <= 34.9 Then
	MsgBox "Obesidade classe I!"
elseif (IMC) >= 35 and (IMC) <= 39.9 Then
	MsgBox "Obesidade classe II"
elseif (IMC) >= 40 then
	MsgBox "Obesidade classe III"
end if 