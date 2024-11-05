set WMI = GetObject ("WinMgmts:")
Dim servicos
Dim impressoras
'servicos
set objs = WMI.InstancesOf("Win32_Service")
for each obj in objs
	servicos = servicos & obj.name & vbcrlf
next
'impressoras
set objs = WMI.InstancesOf("Win32_Printer")
for each obj in objs
	impressoras = impressoras & obj.name & vbcrlf
next
msgbox servicos, 0, "Serviços"
msgbox impressoras, 0, "Impressoras"

'StartService method
'StopService method