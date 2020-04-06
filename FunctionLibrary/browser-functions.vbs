 'To avoid errors due to typos in variable names
 Option Explicit

 'Note: The page object changes with every action done on it. 
 'Remember to always fetch the latest object instance before doing an action on it. (as you see in example in LoginBrowser)
 
'All below functions, although still abstract, contains application/domain specific 'fixed' values. You should adjust them as per your application under test
'for them to work for your AUT.

'For example, in LoginPage function, the 'name' properties (say "usernameField") are specific to 
'application under test. If you are using it in your application, these values will change 
'to your applications webedit and webbutton property names.

'However for all practical purposes, for your given environment once you set these values, these
'values will not change say per different test environments. So you can safely use these as
'an domain-specific-abstract functions. Remember to change the values if you reuse this library.

Function LoginBrowser(strUserName, strPassword)
    ' Get Login page object
    Dim objPageLogin: Set objPageLogin = GetPageObject("Login", "Login")

    'Input User name
    Dim objWebEditUserName: Set objWebEditUserName = GetWebEdit (objPageLogin, "usernameField")
    Call SetWebEdit (objWebEditUserName, strUserName)

    'Input Password
    Dim objWebEditPassword: Set objWebEditPassword = GetWebEdit (objPageLogin, "passwordField")
    Call SetSecureValueInWebEdit (objWebEditPassword, strPassword)

    'Click login button
    Dim objWebButtonLogin: Set objWebButtonLogin = GetWebButton (objPageLogin, "Login" )
    Call ClickWebButton (objWebButtonLogin)
End Function

'Note: Same like login, the below values are fixed for AUT. It could be different for your application but once 
'set, you dont need to parameterise these. Unless your application in different environment has differnt values for these
'in which case it would make perfect sense to parameterise this. 
Function LogoutBrowser()
    ' Get Home page object
    Dim objPageHome: Set objPageHome = GetPageObject("Home", "Home")
    
    ' Click logout button
    Dim objLogoutImage: Set objLogoutImage = GetImage(objPageHome, "Logout", "//TABLE[@id='globalHeaderID']/TBODY[1]/TR[1]/TD[13]/A[1]/DIV[1]/IMG[1]")
    Call ClickWebButton (objLogoutImage)
End Function

Function LoginTestEnvironment(pathConfigXML, testEnv)
	'Get the browser type to navigate to from this config.xml file 
	Dim objXMLTestEnv: Set objXMLTestEnv = GetTestEnvConfigurationObject(pathConfigXML, testEnv)

	'Close all open browser instances of this browser type (say "IE") for test robustness
	Dim strBrowserName: strBrowserName = GetXMLChildNodeValue(objXMLTestEnv, "BrowserName")
	Dim browsersEXEFileName: browsersEXEFileName = GetBrowsersEXEName(strBrowserName)
	Call CloseAllBrowserInstances(browsersEXEFileName)

	'Launch browser and navigate to url of your choice
	Dim strBrowserURL: strBrowserURL = GetXMLChildNodeValue(objXMLTestEnv, "BrowserURL")
	Call LaunchBrowserAndGoToURL(strBrowserName, strBrowserURL)

	'Ensure that page now exists and is fully synced
	' Get Login page object and sync
	Dim objPageLogin: Set objPageLogin = GetPageObject("Login", "Login")
	Call SyncPage(objPageLogin)

	'Login to the test application
	Dim strUserName: strUserName = GetXMLChildNodeValue(objXMLTestEnv, "UserName")
	Dim strPassword: strPassword = GetXMLChildNodeValue(objXMLTestEnv, "EncodedPassword")
	Call LoginBrowser(strUserName, strPassword)

	'SyncPage objPageHome
	' Get Home page object and sync
	Dim objPageHome: Set objPageHome = GetPageObject("Home", "Home")
	Call SyncPage(objPageHome)
End Function