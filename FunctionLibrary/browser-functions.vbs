 'To avoid errors due to typos in variable names
 Option Explicit

 'Note: 
	'1: Browser functions use 'browser-objects' and 'browser-actions' to create steps for creating functions.
		'In general, you will see a pattern as below:
		' Get object
		' Act on it.
	'2: The page object changes with every action done on it. Thus, remember to always fetch the latest object instance before doing an action on it.
	'3: All below functions, although still abstract, contains application/domain specific 'fixed' values. Thus:
		'a) You should adjust them as per your application under test for them to work for your AUT.
		'b) For example, in LoginPage function, the 'name' properties (say "usernameField") are specific to 'application under test. 
		'   If you are using it in "your" application, I expect these values to change for your applications webedit and webbutton property names.
		'c) However once you have changed it, for all practical purposes, for your given environment, these values will not change say per different test environments. 
		'   So you can safely use these as an domain-specific-abstract functions. 
		'd) So again: Remember to change the values if you reuse this library.

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

Function LogoutBrowser() 
 
	' Get Home page object
    Dim objPageHome: Set objPageHome = GetPageObject("Home", "Home")
    
    ' Click logout button
    Dim objLogoutImage: Set objLogoutImage = GetImage(objPageHome, "Logout", "//TABLE[@id='globalHeaderID']/TBODY[1]/TR[1]/TD[13]/A[1]/DIV[1]/IMG[1]")
    Call ClickWebButton (objLogoutImage)
	
End Function

Function LoginTestEnvironment(objXMLTestEnv, objXMLUser)

	'Launch browser and navigate to url of your choice
	Dim strBrowserName: strBrowserName = GetXMLChildNodeValue(objXMLTestEnv, "BrowserName")
	Dim strBrowserURL: strBrowserURL = GetXMLChildNodeValue(objXMLTestEnv, "BrowserURL")
	Call LaunchBrowserAndGoToURL(strBrowserName, strBrowserURL)

	'Ensure that page now exists and is fully synced
	' Get Login page object and sync
	Dim objPageLogin: Set objPageLogin = GetPageObject("Login", "Login")
	Call SyncPage(objPageLogin)

	'Login to the test application
	Dim strUserName: strUserName = GetXMLChildNodeValue(objXMLUser, "UserName")
	Dim strPassword: strPassword = GetXMLChildNodeValue(objXMLUser, "EncodedPassword")
	Call LoginBrowser(strUserName, strPassword)

	'SyncPage objPageHome
	' Get Home page object and sync
	Dim objPageHome: Set objPageHome = GetPageObject("Home", "Home")
	Call SyncPage(objPageHome)
	
End Function

Function CloseTestBrowsers(objXMLTestEnv)

	'Close all open browser instances of this browser type (say "IE") for test robustness
	Dim strBrowserName: strBrowserName = GetXMLChildNodeValue(objXMLTestEnv, "BrowserName")
	Dim browsersEXEFileName: browsersEXEFileName = GetBrowsersEXEName(strBrowserName)
	Call CloseAllBrowserInstances(browsersEXEFileName)	
	
End Function

'Set up
Function SetUp()   
	
	'Get test environment configuration
	Dim objXMLTestEnv: Set objXMLTestEnv = GetTestEnvConfigurationObject()

	'Close test browsers
	Call CloseTestBrowsers(objXMLTestEnv)

	'Get user configuration 
	Dim objXMLUser: Set objXMLUser = GetUserConfigurationObject()

	'Login to Test environment
	Call LoginTestEnvironment(objXMLTestEnv, objXMLUser)
	
End Function

'Tear Down
Function TearDown()   
	
	'Logout from browser
	Call LogoutBrowser()
	
	'Get test environment configuration
	Dim objXMLTestEnv: Set objXMLTestEnv = GetTestEnvConfigurationObject()

	'Close test browsers
	Call CloseTestBrowsers(objXMLTestEnv)
	
End Function

' NavigateToOracleForms via a general module (System administrator)	
' Then in tests (not here), switch to the responsibility that you want to test.
Function NavigateToOracleForms()   
	
	' Get Home page object and sync
	Dim objPageHome: Set objPageHome = GetPageObject("Home", "Home")
	Call SyncPage(objPageHome)

	'Now Navigate to 'System Administrator'
	Dim objLinkSystemAdministrator: Set objLinkSystemAdministrator = GetLink2(GetPageObject("Home", "Home"), "System Administrator")
	Call ClickLink(objLinkSystemAdministrator)

	'Open 'User Monitor' Link
	Dim objLinkUserMonitor: Set objLinkUserMonitor = GetLink2(GetPageObject("Home", "Home"), "User Monitor")
	Call ClickLink(objLinkUserMonitor)

	'Wait for Oracle page to sync
	Dim objPageOracle: Set objPageOracle = GetPageObject("Oracle E-Business Suite R12", "Oracle E-Business Suite R12")
	Call SyncPage(objPageOracle)

	'msgbox "Web parts done. Now Oracle application part starts..."
	
End Function