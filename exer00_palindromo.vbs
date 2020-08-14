dim palavra
dim palavraInvertida

palavra = InputBox("Insira a palavra:")
palavraInvertida = StrReverse(palavra)

if palavra = palavraInvertida then

	WScript.Echo "É palíndromo!"

else

	WScript.Echo "Não é palíndromo!"

end if