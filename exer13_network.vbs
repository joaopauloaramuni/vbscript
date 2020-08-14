Set Rede = WScript.CreateObject("WScript.Network")
PropertyInfo = "Domínio do usuário" & vbTab & "= " & Rede.UserDomain & _
vbCrLf & "Computador" & vbTab & "= " & Rede.ComputerName & _
vbCrLf & "Nome do usuário" & vbTab & "= " & Rede.UserName & vbCrLf
MsgBox(PropertyInfo)
