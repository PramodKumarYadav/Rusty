 'To avoid errors due to typos in variable names
 Option Explicit
 
 ' List of all 18 oracle objects: https://admhelp.microfocus.com/uft/en/14.50-14.53/UFT_Help/Subsystems/FunctionReference/Subsystems/OMRHelp/Content/Oracle/ORACLEPACKAGELib_P.html?TocPath=Object%20Model%20Reference%20for%20GUI%20Testing%7COracle%7C_____0

' Example Recording: OracleNotification("Caution").Approve
Function GetOracleNotification(title)
	
	'Set object based on the parent object and property title
	Dim objOracleNotification: Set objOracleNotification = OracleNotification("title:="&title)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleNotification, title, "OracleNotification" 
	
	'Assign this object to function
	Set GetOracleNotification = objOracleNotification
	
	'Now release this object memory
	Set objOracleNotification = Nothing
	
End Function

Function CheckIfObjectExistsAndIsEnabled (objToCheck, objPropertyValue, objType )
	
	'Check if Oracle object exists
	CheckIfOracleObjectExists objToCheck, objPropertyValue, objType
	
	'Check if object is enabled
	CheckIfObjectIsEnabled objToCheck, objPropertyValue, objType
	
End Function

Function CheckIfOracleObjectExists(objToCheck, objPropertyValue, objType)
	
	'Set status to fail and exit test if object doesnt exist
	' Do not add status for pass (using micDone), since UFT by default will add it for below check i.e. objToCheck.exist (adding it will result in duplicate statements)	
	
	' Report the result	
	Dim testStep: testStep = "Find object type ["&objType&"] with property value ["&objPropertyValue&"]"
	If objToCheck.exist Then		
		Call SetResultRecordForTestReport(testStep, "Found","Found", "Pass")		
	Else
		Call SetResultRecordForTestReport(testStep, "Not Found","Found", "Fail")
		Call CreateReportAndExitTests()
	End if
	
End Function

Function CheckIfObjectIsEnabled (objToCheck, objPropertyValue, objType)
	
	' Check if object is enabled (sometimes the object exists but is not enabled and thus the operations will fail).		
	' Do not add status for pass (using micDone), since UFT by default will add it for below check i.e. objToCheck.WaitProperty (adding it will result in duplicate statements)	
	
	' Report the result	
	Dim testStep: testStep = "Check if Object type ["&objType&"] with property value ["&objPropertyValue&"] is enabled"
	If objToCheck.WaitProperty ("enabled", "True", 40000) Then			
		Call SetResultRecordForTestReport(testStep, "Enabled","Enabled", "Pass")		
	Else
		Call SetResultRecordForTestReport(testStep, "Not Enabled","Enabled", "Fail")
		Call CreateReportAndExitTests()
	End if
End Function

' Example Recording: OracleFormWindow("Navigator").OracleTabbedRegion("Functions")
Function GetOracleFormWindow(title)
	
	'Set object based on the parent object and property title
	Dim objOracleFormWindow: Set objOracleFormWindow = OracleFormWindow("title:="&title)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleFormWindow, title, "OracleFormWindow" 
	
	'Assign this object to function
	Set GetOracleFormWindow = objOracleFormWindow
	
	'Now release this object memory
	Set objOracleFormWindow = Nothing
	
End Function

' Example Recording: OracleListOfValues("Responsibilities").Select "Test User"
Function GetOracleListOfValues(title)
	
	'Set object based on the parent object and property title
	Dim objOracleListOfValues: Set objOracleListOfValues = OracleListOfValues("title:="&title)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleListOfValues, title, "OracleListOfValues" 
	
	'Assign this object to function
	Set GetOracleListOfValues = objOracleListOfValues
	
	'Now release this object memory
	Set objOracleListOfValues = Nothing
	
End Function

' Example Recording: OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("List").Select "Party Reports"
Function GetOracleTabbedRegion(objParent, label)
	
	'Set object based on the parent object and property label
	Dim objOracleTabbedRegion: Set objOracleTabbedRegion = objParent.OracleTabbedRegion("label:="&label)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleTabbedRegion, label, "OracleTabbedRegion" 
	
	'Assign this object to function
	Set GetOracleTabbedRegion = objOracleTabbedRegion
	
	'Now release this object memory
	Set objOracleTabbedRegion = Nothing
	
End Function

' Example Recording: OracleFormWindow("Navigator").OracleTable("Functions")
Function GetOracleTable(objParent, columns)
	
	'Set object based on the parent object and property label
	Dim objOracleTable: Set objOracleTable = objParent.OracleTable("columns:="&columns)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleTable, columns, "OracleTable" 
	
	'Assign this object to function
	Set GetOracleTable = objOracleTable
	
	'Now release this object memory
	Set objOracleTable = Nothing
	
End Function

'Note: If you are recording, depending on what items are open the property 'selected index' can change. Also, You can have an item with same name in two lists. 
'By using a "parent object and calling with name 'selected item'", seems like a more robust and clean appraoch. 

' Example Recording: OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("List")
Function GetOracleList(objParent, strDescription)
	
	'Set object based on the parent object and property selected_item
	Dim objOracleList: Set objOracleList = objParent.OracleList("description:="&strDescription)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleList, strDescription, "OracleList" 
	
	'Assign this object to function
	Set GetOracleList = objOracleList
	
	'Now release this object memory
	Set objOracleList = Nothing
	
End Function

' Example Recording: OracleFormWindow("Invoices").OracleTextField("Partner/Some id field")
Function GetOracleTextField(objParent, strDescription)
	
	'Set object based on the parent object and property selected_item
	Dim objOracleTextField: Set objOracleTextField = objParent.OracleTextField("description:="&strDescription)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleTextField, strDescription, "OracleTextField" 
	
	'Assign this object to function
	Set GetOracleTextField = objOracleTextField
	
	'Now release this object memory
	Set objOracleTextField = Nothing
	
End Function

' Example Recording: OracleFormWindow("Invoices").OracleTextField("Partner/Some id field")
Function GetOracleTextFieldUsingDescAndToolTip(objParent, strDescription, tooltip)
	
	'Set object based on the parent object and property selected_item
	Dim objOracleTextField: Set objOracleTextField = objParent.OracleTextField("description:="&strDescription, "tooltip:="&tooltip)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleTextField, strDescription, "OracleTextField" 
	
	'Assign this object to function
	Set GetOracleTextFieldUsingDescAndToolTip = objOracleTextField
	
	'Now release this object memory
	Set objOracleTextField = Nothing
	
End Function

' Example Recording: OracleFormWindow("Find Invoices").OracleButton("Find").Click
Function GetOracleButton(objParent, strDescription)
	
	'Set object based on the parent object and property selected_item
	Dim objOracleButton: Set objOracleButton = objParent.OracleButton("description:="&strDescription)

	'Check and Continue only if the object exists and is enabled
	CheckIfObjectExistsAndIsEnabled objOracleButton, strDescription, "OracleButton" 
	
	'Assign this object to function
	Set GetOracleButton = objOracleButton
	
	'Now release this object memory
	Set objOracleButton = Nothing
	
End Function
