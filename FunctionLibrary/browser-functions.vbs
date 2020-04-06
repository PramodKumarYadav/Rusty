 'To avoid errors due to typos in variable names
 Option Explicit

'All below functions, although still abstract, contains application specific 'fixed' values.

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