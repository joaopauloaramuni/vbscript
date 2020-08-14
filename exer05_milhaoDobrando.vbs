' Suponha que você comece com 1 real e vá dobrando seu dinheiro todos os dias. Quantos dias iriam levar para
' você juntar 1 milhão de reais?

 
Option Explicit                  ' must declare every variables before use

const MAXAMOUNT = 1000000
dim dobro                        ' current amount of money
dim sum                          ' total sum of accumulated money
dim days                         ' number of days

' initialize 
dobro = 1
sum = 0
days = 0

'Outra forma de fazer loops
do until sum > MAXAMOUNT 

   sum = sum + dobro  '1 3 7 15 31...
   dobro = 2 * dobro '2 4 8 16..
   days = days + 1 	   
   
   WScript.Echo "Dia: " & days & _
   vbCrLf  & "Total no dia: " & sum   

loop

MsgBox "Começando com 1 real e dobrando seu dinheiro todos os dias, levaria " _
       & days & " dias para ter R$ " & sum & " reais." 