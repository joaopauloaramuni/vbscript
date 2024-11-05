' Verifique se a porta foi criada:
' Firewall do Windows -> Configurações Avançadas -> Regras de Entrada -> Porta de Teste

Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set objPort = CreateObject("HNetCfg.FwOpenPort")
objPort.Port = 9999
objPort.Name = "Porta de Teste"
' Habilitar / Desabilitar porta
objPort.Enabled = TRUE

' Adicionar porta
Set colPorts = objPolicy.GloballyOpenPorts
colPorts.Add(objPort)
