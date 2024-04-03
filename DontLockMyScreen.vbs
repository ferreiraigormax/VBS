Dim objResult

set objShell = WScript.CreateObject("WScript.Shell")
Do While True
	objResult = objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
	Wscript.sleep(110000)
Loop
