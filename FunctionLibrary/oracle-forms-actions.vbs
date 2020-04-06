 'To avoid errors due to typos in variable names
 Option Explicit
 
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

' ------------------
' OracleListItem actions
' ------------------
' Example Recording: OracleFormWindow("Navigator").OracleTabbedRegion("My Functions").OracleList("Some Function List").Activate "+  Navigate"
Function ActivateOracleListItem(object, selected_item)	
	
	'Activate this list item 
	object.Activate(selected_item)	
	
End Function

' ------------------
' OracleTextField actions
' ------------------
' Example Recording: OracleFormWindow("Invoices").OracleTextField("Trading Partner/Some id field").Enter "someid"
Function SetOracleTextField(object, value)
	
	'Set the text field with given value
	object.Enter(value)
	
End Function

' ------------------
' OracleButton actions
' ------------------
' Example Recording: OracleFormWindow("Find All Invoices").OracleButton("Find").Click
Function ClickOracleButton(object)
	
	' Click the button
	object.Click
	
End Function
