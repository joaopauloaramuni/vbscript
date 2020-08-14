Function Fibonacci(N)
  If N < 2 Then
    Fibonacci = N
  Else
    Fibonacci = Fibonacci(N - 1) + Fibonacci(N - 2)
  End If
End Function

For i = 1 To 20
  res = res & Fibonacci(i) & ", "
Next
WScript.Echo (res & "...")