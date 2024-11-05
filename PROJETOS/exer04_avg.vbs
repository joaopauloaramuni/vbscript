'Cálculo da Média
 
Option Explicit                  ' Obrigo a declarar todas as variáveis antes de usar

dim inputNumber                  ' input de entrada dos números
dim count                        ' contador de entradas do usuário
dim sum                          ' soma dos números inseridos pelo usuário
dim avg                          ' média dos números inseridos pelo usuário

' Inicilizar variáveis para evitar erros
count = 0
sum = 0
avg = 0

' Para todas as entradas de usuário, contamos e somamos os valores inseridos
do
   inputNumber = InputBox("Entre com um número. Para sair digite um número negativo.")

   ' Verifica se o número inserido é númerico
   if not IsNumeric(inputNumber) then
      MsgBox "Você deve inserir um valor numérico, tente de novo.", vbOKOnly, "Dados Inválidos"
   else
      if inputNumber < 0 then
         exit do
      end if
      sum = sum + inputNumber
      count = count + 1
   end if
loop

' Calcular o valor da média (média = soma dos dados inseridos dividido pelo número de dados inseridos)
' Desafio: Considerando as boas práticas de programação, o cálculo da média deve ser feito em uma procedure
' separadamente. Tente passar isso para um Sub e/ou para uma Function.
if count <> 0 then
   avg = sum/count
else
   avg = 0
end if
MsgBox "Número de itens inseridos:  " & count & "  Média:  " & avg,  vbOKOnly, "Resultado"