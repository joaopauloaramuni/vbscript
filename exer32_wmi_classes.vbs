'Script para retornar todas as classes wmi disponíveis no sistema
strComputer  = "."
strNamespace = "root\cimv2"

Set oWMI = GetObject("winmgmts:\\" & strComputer & "\" & strNamespace)
Set colClasses = oWMI.ExecQuery("SELECT * FROM meta_class") 

For Each oClass in colClasses

  For Each oQualifier In oClass.Qualifiers_
    strQualName = LCase(oQualifier.Name)

    If strQualName = "dynamic" OR strQualName = "static" Then
      If oClass.Methods_.Count > 0 Then
        WScript.Echo oClass.Path_.Class
      End If
    End If

  Next
Next



















