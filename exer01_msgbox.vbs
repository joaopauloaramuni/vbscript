'Existem duas maneiras de se executar um aquivo com extensão .vbs no WSH
'1 - Através do comando cscript arquivo.vbs no cmd
'2 - Clicando duas vezes sobre ele

Msgbox("Exemplo de MsgBox")

Dim result
result = MsgBox ("Beleza?", vbYesNo, "Sim/Nao")
Select Case result
    Case vbYes
		MsgBox("Eu também.")
    Case vbNo
        MsgBox("Relaxa...curta a aula!")
End Select