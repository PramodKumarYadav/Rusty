 'To avoid errors due to typos in variable names
 Option Explicit

'Navigate to the URL 
Function NavigateToURL(objBrowser, URL)
    
    objBrowser.navigate(URL)
	
End Function

'Sync this particular page
Function SyncPage(objPage)
	
    objPage.Sync
	
End Function

'Set value in webedit 
Function SetWebEdit (objWebEdit, strValue)
	
    objWebEdit.Set strValue
	
End Function

'Set secure value in webedit 
Function SetSecureValueInWebEdit (objWebEdit, strValue)
	
    objWebEdit.SetSecure strValue
	
End Function

'Click the button
Function ClickWebButton (objWebButton)
	
    objWebButton.Click
	
End Function