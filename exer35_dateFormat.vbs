' Trabalhando com datas no VBScript
' A manipulação de datas é um dos maiores desafios da programação distribuída
' Devemos tomar cuidado, pois podemos estar recuperando a data do servidor, que pode estar em outro país
' e por esse motivo apresentar um GMT diferente (horas a mais ou a menos), e toda nossa lógica pode ser comprometida

Dim data
data = now

WScript.Echo "Data completa: " & now

WScript.Echo "Data por extenso: " & FormatDateTime (now, 1)

WScript.Echo "Apenas data: " & FormatDateTime (now, 2)

WScript.Echo "Apenas hora: " & FormatDateTime (now, 4)

WScript.Echo "Ano: " & year(data)

WScript.Echo "Mês: " & month(data)

WScript.Echo "Dia: " & day(data)

WScript.Echo "Hora: " & hour(data)

WScript.Echo "Minuto: " & minute(data)

WScript.Echo "Segundo: " & second(data)