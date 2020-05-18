'To avoid errors due to typos in variable names
Option Explicit

' To launch powershell and run a powershell script
Function RunPowershellScript(pathPS1Script)
	Set objShell = CreateObject("Wscript.shell")
	' Wrap the path in double quotes so that if there is any space in folder names, it still works okay.
	dim statement: statement = "pwsh -noexit -file """ & pathPS1Script & """"
	objShell.run(statement)
End Function