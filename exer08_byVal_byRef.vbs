'Demonstrando a diferença entre o tipo byVal e o tipo ByRef

option explicit

dim a, b, c, d
a = 2
b = 3
c = 2
d = 3

call byValSub(a, b)
call byRefSub(c, d)

MsgBox "a = " & a & "  b = " & b & "  c = " & c & "  d = " & d

'----------------------------------------------------------------------------
' byValSub
' Os parâmetros passados são copiados e, caso sejam mudamos, são mudados localmente na procedure,
' quando a procedure termina, estas variáveis cópias morrem, sem alterar os valores das variáveis originais.
SUB byValSub(ByVal num1, ByVal num2)
   num1 = 10
   num2 = 20
end SUB

'----------------------------------------------------------------------------
' byRefSub
' Os parâmetros são passados como referência na memória e, caso sejam mudados, são mudados globalmente no programa,
' ou seja, são alterados os valores das variáveis originais.
SUB byRefSub(ByRef num1, ByRef num2)
   num1 = 10
   num2 = 20
end SUB

'Em VBScript temos dois tipos de procedures:

' Sub procedure
' Function procedure

' As procedures do tipo Sub não retornam nenhum valor
' As procedures do tipo Function retornam valor

' Você pode ver um exemplo de retorno no exer10_fibonacci.sh

' Ou como abaixo:

' Function minhaFuncao()
  'Comandos...
  'minhaFuncao = Algum valor de retorno...pode ser do tipo inteiro ou string por exemplo,
  'não é necessário explicitar o tipo do retorno, ex:
  'minhaFuncao = 0
  'ou minhaFuncao = "Olá Mundo"
' End Function