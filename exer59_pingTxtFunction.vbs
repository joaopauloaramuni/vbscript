Function Ping( myHostName )
' A função retorna Verdadeiro se o host especificado puder ser pingado com sucesso.
' O parâmetro myHostName pode ser o nome de um computador ou o IP address.
' A classe Win32_PingStatus usada nessa função requer Windows XP ou superior.

    ' Variáveis
    Dim colPingResults, objPingResult, strQuery

    ' Define a query WMI
    strQuery = "SELECT * FROM Win32_PingStatus WHERE Address = '" & myHostName & "'"

    ' Executa a query WMI
    Set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery( strQuery )

    For Each objPingResult In colPingResults
        If Not IsObject( objPingResult ) Then
            Ping = False
        ElseIf objPingResult.StatusCode = 0 Then
            Ping = True
        Else
            Ping = False
        End If
    Next

    Set colPingResults = Nothing
End Function

'Script VBS – Pingar IP's lidos de um arquivo .txt
'Variável que representa um ip
dim ip
'Constante que será passada como parâmetro para o método openTextFile da classe FSO
const ForReading = 1

'fso - obj da classe FileSystemObject
set fso = CreateObject("Scripting.FileSystemObject")
'fileTxt - obj que representa o arquivo texto .txt 
'O arquivo texto será retornado através do método openTextFile da classe FSO
set fileTxt = fso.openTextFile("U:\ips.txt",ForReading)

'Loop para ler o arquivo txt até o final
Do until fileTxt.AtEndOfStream	
	'Leitura da linha do txt que possui cada ip
	ip = fileTxt.readLine
	'Execução do comando ping
	WScript.Echo "IP: " & ip & VbCrLf & "Pingado: " & ping(ip) & VbCrLf
loop