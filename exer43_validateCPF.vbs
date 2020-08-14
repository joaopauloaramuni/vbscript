function ValidateCPF(cpf)
    Dim multiplic1, multiplic2
    multiplic1=Array(10, 9, 8, 7, 6, 5, 4, 3, 2)
    multiplic2=Array(11, 10, 9, 8, 7, 6, 5, 4, 3, 2 )
    Dim tempCpf,digit,sum,remainder,i,RegXP
    cpf = Trim(cpf)
    cpf = Replace(cpf,".", "")
    cpf = Replace(cpf,"-", "")
    if (Len(cpf) <> 11) Then
        ValidateCPF = false
    else
        tempCpf = Left (cpf, 9)
        sum = 0

        Dim intCounter
        Dim intLen 
        Dim arrChars()

        intLen = Len(tempCpf)-1
        redim arrChars(intLen)

        For intCounter = 0 to intLen
            arrChars(intCounter) = Mid(tempCpf, intCounter + 1,1)
        Next

        i=0
        For i = 0 to 8
            sum =sum + CInt(arrChars(i)) * multiplic1(i)
        Next

        remainder = sum Mod 11
        If (remainder < 2) Then
            remainder = 0
        else
            remainder = 11 - remainder
        End If

        digit = CStr(remainder)
        tempCpf = tempCpf & digit
        sum = 0

        intLen = Len(tempCpf)-1
        redim arrChars(intLen)
        intCounter= 0
        For intCounter = 0 to intLen
            arrChars(intCounter) = Mid(tempCpf, intCounter + 1,1)
        Next
        i=0
        For i = 0 to 9
            sum =sum + CInt(arrChars(i)) * multiplic2(i)
        Next        
        remainder = sum Mod 11

        If (remainder < 2) Then
            remainder = 0
        else
            remainder = 11 - remainder
        End If      
        digit = digit & CStr(remainder)

        Set RegXP=New RegExp
            RegXP.IgnoreCase=1
            RegXP.Pattern=digit & "$"

        If RegXP.test(cpf) Then 
            ValidateCPF = true
        else
            ValidateCPF = false
        end if
    end if
end Function

if ValidateCPF(InputBox("Digite o seu CPF.","Informe.")) = False then
    msgbox "CPF Inválido."
else
    msgbox "CPF Válido."
end if