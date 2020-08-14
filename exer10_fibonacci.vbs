'Funcao Recursiva
Function Fibonacci(N)
  If N < 2 Then
    Fibonacci = N
  Else
    Fibonacci = Fibonacci(N - 1) + Fibonacci(N - 2)
  End If
End Function

n = InputBox("Digite um número:")

res = "Fibonacci de " & n & " é: " & Fibonacci(n)

WScript.Echo (res)

' Outro modo de fazer
'Function Fibonacci(N)
	'if N = 0 Then
		'Fibonacci = 0
	'elseif n = 1 Then
		'Fibonacci = 1
	'else
	  'Fibonacci = Fibonacci(N - 1) + Fibonacci(N - 2)
	'End if
'End Function
