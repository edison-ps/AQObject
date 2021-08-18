Dim agora

agora = Now
WScript.Sleep 3571
WScript.Echo "Tempo total de execucao: " + CStr(TimeSerial (0, 0, DateDiff("s", agora, Now)))
'WScript.Echo FormatDateTime(DateDiff("m", agora, Now), 3)  

For i = 1 to 10

	WScript.Echo Now
	WScript.Sleep 1000

Next

