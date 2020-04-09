 'To avoid errors due to typos in variable names
 Option Explicit

 'Note: These functions cannot be called from UFT. You can use them in a bat file outside UFT to launch UFT. Not from UFT.
Function GetUFTObject()
	
	'Open QTP/UFT
	Set objQTP = CreateObject("QuickTest.Application")

	'Assign this object to function
	Set OpenUFT = objQTP
	
	'Now release this object memory
	Set objQTP = Nothing
	
End Function

Function OpenUFT(objQTP)
	
	'Launch UFT and make it visible
	objQTP.Launch
	objQTP.Visible = True
	
End Function