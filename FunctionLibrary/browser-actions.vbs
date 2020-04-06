 'To avoid errors due to typos in variable names
 Option Explicit

  ' Get browsersEXEName for a browser type
Function GetBrowsersEXEName(strBrowserType)	
    Dim exeFileName	
    Select Case UCase(strBrowserType)
        Case "IE", "INTERNET EXPLORER"
            exeFileName = "iexplore.exe"		
        Case "FF", "FIREFOX"
            exeFileName = "firefox.exe"
        Case "GC", "GOOGLE CHROME"
            exeFileName = "chrome.exe"	
        Default 
            exeFileName = "iexplore.exe"
    End Select

    'Return EXE file name
    GetBrowsersEXEName = exeFileName
End Function

'Close all open browser instances of selected browser type. If there are more than one instances open, we will not be able to perform operations on page. 
Function CloseAllBrowserInstances(browsersEXEName)
    SystemUtil.CloseProcessByName(browsersEXEName)
End Function

'Launch browser with this exe file name
Function LaunchBrowser(browsersEXEName)   
    SystemUtil.Run browsersEXEName
End Function

 ' Launch and Navigate to a URL
 Function LaunchBrowserAndGoToURL(strBrowserType, strURLAddress)
    'Get exe file name
    Dim exeFileName: exeFileName = GetBrowsersEXEName(strBrowserType)

    'Launch URL from chosen browser
    SystemUtil.Run exeFileName,strURLAddress
End Function

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

'Click link
Function ClickLink (objLink)
	
    objLink.Click
	
End Function