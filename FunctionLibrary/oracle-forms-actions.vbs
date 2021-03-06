'To avoid errors due to typos in variable names
 Option Explicit
 
' List of actions available for all 18 oracle objects can be found here: https://admhelp.microfocus.com/uft/en/14.50-14.53/UFT_Help/Subsystems/FunctionReference/Subsystems/OMRHelp/Content/Oracle/ORACLEPACKAGELib_P.html?TocPath=Object%20Model%20Reference%20for%20GUI%20Testing%7COracle%7C_____0

' ------------------
' OracleNotification actions
' ------------------
' Example Recording: OracleNotification("Caution").Approve
Function ApproveOracleNotification(object)	
	
	' Approve the notification window
	object.Approve
	
End Function

' ------------------
' OracleFormWindow actions
' ------------------
' Example Recording: OracleFormWindow("Find Invoices").CloseWindow
Function CloseOracleFormWindow(object)	
	
	' Close the form window
	object.CloseWindow
	
End Function

Function SelectMenuOracleFormWindow(object, menuItem)	
	
	' Select a menu item
	object.SelectMenu(menuItem)
	
End Function
' ------------------
' OracleListOfValues actions
' ------------------
' Example Recording: OracleListOfValues("Responsibilities").Select "Test User"
Function SelectAnItemFromOracleListOfValues(object, selected_item)
	
	' Select an item from list
	object.Select(selected_item)
	
End Function

' ------------------
' OracleListItem actions
' ------------------
' Example Recording: OracleFormWindow("Navigator").OracleTabbedRegion("My Functions").OracleList("Some Function List").Activate "+  Navigate"
Function ActivateOracleListItem(object, selected_item)	
	
	'Activate this list item 
	object.Activate(selected_item)	
	
End Function

' ------------------
' OracleTable actions
' ------------------
Function GetFieldValueFromOracleTable(object, RecordNumber,ColumnNr)	
	
	'Return the field value at a record, column nr
	GetFieldValueFromOracleTable = object.GetFieldValue(RecordNumber,ColumnNr)	
	
End Function


' Signature: Enter ([in] Text, [in, optional] WithValidation)
' Description: Enters the specified text into the field.
' Example Recording: OracleFormWindow("Invoices").OracleTextField("Trading Partner/Some id field").Enter "someid"
Function SetOracleTextFieldValue(object, text)
	
	object.Enter(text)
	
End Function


' Signature: GetROProperty ([in] Property, [in, optional] PropertyData)
' Description: Returns the current value of the specified identification property from the object in the application.
' Example: Mostly we will pass this txt as "value" - value being the "property name" of (say) text box "value".
Function GetValueOfSpecifiedPropertyFromObject(object, propertyName)
	
	GetValueOfSpecifiedPropertyFromObject = object.GetROProperty(propertyName)

End Function


' Signature: Click ([in, optional] x, [in, optional] y, [in, optional] BUTTON)
' Description: Clicks the specified location with the specified mouse button.
' Example Recording: OracleFormWindow("Find All Invoices").OracleButton("Find").Click
Function ClickObject(object)
		
	object.Click
	
End Function


' Signature: InvokeSoftkey ([in] Softkey)
' Description: Invokes the specified Oracle softkey.
' Example Recording: OracleFormWindow("Transactions").OracleTextField("Transaction|Source").InvokeSoftkey "ENTER QUERY" (for F11) & "EXECUTE QUERY" for (Ctrl+F11)
Function InvokeSoftkeyOnObject(object, Softkey)	
	
	object.InvokeSoftkey(Softkey)
	
End Function
