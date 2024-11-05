n = InputBox("Digite um número:")

dim f 
f=1 
if n<0 then 
	Msgbox "Número inválido!" 
	elseif n=0 or n=1 then 
	MsgBox "O fatorial do número "&n&" é :"&f 
	else 

	for i=n to 2 step -1 
		f=f*i 
	next 

	MsgBox "O fatorial do número "&n&" é :"&f 

end if 