Set objVoice = CreateObject("SAPI.SpVoice")

For Each strVoice in objVoice.GetVoices
	'Mostrar as vozes disponíveis
    Wscript.Echo strVoice.GetDescription
Next

'Voz da Maria - Português
Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Maria Desktop").Item(0)
objVoice.Speak "Oi eu sou a Maria."

'Voz da Zira - Inglês
Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Zira Desktop").Item(0)
objVoice.Speak "Hi i am Zira"