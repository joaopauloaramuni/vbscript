Dim strPath, strScreenW, strScreenH, strTitle, boolValid
Dim AFDTransaction, AFDColumns(), AFNColumns(), AFNRows(), AFNRow, AFDRow, strPathAFD, strInitial, strFinal

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Const COLUMN = "AB:"
Const INITIAL = "i:"
Const FINAL = "f:"

strPath = Mid(wscript.ScriptFullName,1,InStr(wscript.ScriptFullName,wscript.ScriptName) - 1)

boolValid = 0
RowsIndex = 0
AFDTransaction = Array()

strPathAFD = strPath & "\afd.txt"

Set FSO = CreateObject("Scripting.FileSystemObject")

Set AFN = FSO.OpenTextFile(strPath & "\afn.txt", ForReading)
Set AFD = FSO.OpenTextFile(strPathAFD, ForWriting, True)

Do Until AFN.AtEndOfStream

	
	AFNLine = Split(AFN.Readline , " ")

	If Not IsArray(AFNLine) Or UBound(AFNLine) < 0 Then
		MsgBox "Padrão do Arquivo ã inválido ou estã vazio!", vbCritical, "Formato Inválido"
		WScript.Quit
	End If
	
	If AFNLine(0) = COLUMN Then
		boolValid = boolValid + 1
		
		ReDim AFNColumns(Ubound(AFNLine)) 
		
		AFD.Write COLUMN
		
		For i = 1 to Ubound(AFNLine)
			AFD.Write " "
			AFD.Write AFNLine(i)
			AFNColumns((i-1)) = AFNLine(i)
		Next

		AFD.Write vbCrLf
		
	ElseIf AFNLine(0) = INITIAL Then
		AFD.WriteLine INITIAL & " " & AFNLine(1)
		boolValid = boolValid + 1
	ElseIf AFNLine(0) = FINAL Then
		AFD.WriteLine FINAL
		boolValid = boolValid + 1
	Else
		
		If boolValid < 3 Then
			MsgBox "Padrão do Arquivo inválido!", vbCritical, "Formato Inválido"
			WScript.Quit
		End If

		boolColumn = False
		AFNRowCol = Empty 'AFNLine(1) & " "
		AFNRow = AFNRowCol
		
		For i = 0 to Ubound(AFNLine)
		
			If boolColumn = True Then
				If Not AFNRow = AFNRowCol Then	
					AFNRow = AFNRow & "," & AFNLine(i)
				Else
					AFNRow = AFNRow & AFNLine(i)
				End If
			End If
		
			For j = 0 to Ubound(AFNColumns)
				If AFNLine(i) = AFNColumns(j) Then
					boolColumn = True
				End If				
			Next
						
		Next  
		
		If AFNRow = AFNRowCol Then
			AFNRow = Empty
		End If
				
		If Not IsEmpty(AFNRow) Then	

			ReDim Preserve AFNRows(RowsIndex)
			
			AFNRows(RowsIndex) = AFNRow

			RowsIndex = RowsIndex + 1

		End If 

	End If
	
Loop

RowsIndex = 0

For i = 0 to Ubound(AFNRows) -1
	
	For j = 0 to UBound(AFNColumns) -1
	
			AFDRow = Empty

			If AFNRows(i) = AFNRows(j) Then
				AFDRow = AFNRows(j)
				'AFD.WriteLine AFNRows(i) & " " &  AFNColumns(j) & " " & AFDRow
			ElseIf i = 1 Then
				AFDRow = AFNRows(j) & "," & AFNRows(i+1)
				'AFD.WriteLine AFNRows(i) & " " &  AFNColumns(j) & " " & AFDRow
			ElseIf i = 2 And j = 1 Then
				AFDRow = AFNRows(j) & "," & AFNRows(i+1)
				'AFD.WriteLine AFNRows(i) & " " &  AFNColumns(j) & " " & AFDRow
			ElseIf i = 3 And j = 0 Then
				AFDRow = AFNRows(j) & "," & AFNRows(i-1) & "," & AFNRows(i+1)
				'AFD.WriteLine AFNRows(i) & " " &  AFNColumns(j) & " " & AFDRow
			Else
				AFDRow = AFNRows(j)
				'AFD.WriteLine AFNRows(i) & " " &  AFNColumns(j) & " " & AFDRow
			End If
		
			'AFD.WriteLine AFDRow
			
			ReDim Preserve AFDColumns(RowsIndex)

			AFDColumns(RowsIndex) = AFDRow
			
			RowsIndex = RowsIndex + 1
		
	Next
		
Next 

Call doTransaction()

RowsIndex = 0

For i = 0 to UBound(AFDTransaction)
	For j = 0 to UBound(AFNColumns) - 1
			
		AFD.WriteLine AFDTransaction(i) & " " &  AFNColumns(j) & " " & AFDColumns(RowsIndex)
		
		If UBound(AFDColumns) > RowsIndex Then
			RowsIndex = RowsIndex + 1
		Else
			RowsIndex = 4
		End If
		
	Next
Next 

AFD.Close
AFN.Close

Call ReplaceFinal()

Sub ReplaceFinal()

	Set AFDRead = FSO.OpenTextFile(strPathAFD, ForReading)
	
	textFinal = AFDRead.ReadAll
	AFDRead.Close
	
	strNewFinal = Replace(textFinal, FINAL, FINAL & " " & AFDColumns(UBound(AFDColumns) -1))
	
	Set AFDWrite = FSO.OpenTextFile(strPathAFD, ForWriting, True)
	AFDWrite.WriteLine strNewFinal
	AFDWrite.Close
	
End Sub

Sub doTransaction()

	For i = 0 to UBound(AFDColumns) - 1
	
		BoolExist = False
		SizeTrans = (UBound(AFDTransaction) - 1) 
	
		For j = 0 to SizeTrans
			If AFDTransaction(j) = AFDColumns(i) Then
				BoolExist = True
			End If
		Next
		
		If Not BoolExist Then
			
			If IsArray(AFDTransaction) Then
				SizeTrans = (UBound(AFDTransaction) + 1)
			Else
				SizeTrans = 1
			End If 
		
			ReDim Preserve AFDTransaction(SizeTrans)
			AFDTransaction(SizeTrans) = AFDColumns(i)
			
		End If
	
	Next
		
End Sub



WScript.Echo "Conversão finalizada!"
WScript.Quit

