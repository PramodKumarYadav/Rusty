
'To avoid errors due to typos in variable names
Option Explicit

'Below is a demo script to find Invoices. This is an example to show how you can create funtions using objects and actions to make your function Tests.
	' Common approach is "First: Get object ; Second: Act on the object

Function FindInvoiceDemo(recordSet)   
	
	' Now entering into Oracle applictions
	'OracleNotification("Error").Approve
	Dim objOracleNotification: Set objOracleNotification = GetOracleNotification("Error")
	Call ApproveOracleNotification(objOracleNotification)

	'Navigate to the Invoices form
	'OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Activate "+  Navigate"
	Dim objOracleFormWindow: Set objOracleFormWindow = GetOracleFormWindow("Navigator")
		Dim objOracleTabbedRegion: Set objOracleTabbedRegion = GetOracleTabbedRegion(objOracleFormWindow, "Functions")
			Dim objGetOracleList: Set objGetOracleList = GetOracleList(objOracleTabbedRegion, "Function List")
			Call ActivateOracleListItem (objGetOracleList, "+  Navigate List")
			Call ActivateOracleListItem (objGetOracleList, "   +  Invoices List")
			Call ActivateOracleListItem (objGetOracleList, "      +  Inquiry")
	
	Do Until recordSet.EOF		
		
		' Open the Invoices window to search a record.
		Call ActivateOracleListItem (objGetOracleList, "             Invoices")
		
		' Get Form Window
		Dim objFindInvoicesFormWindow: Set objFindInvoicesFormWindow = GetOracleFormWindow("Find Invoices")
		
		' Set text field with Trading Partner: Name
		Dim objOracleTextFieldName: Set objOracleTextFieldName = GetOracleTextField(objFindInvoicesFormWindow, "PartnerName")
		Call SetOracleTextFieldValue(objOracleTextFieldName, recordSet.Fields(0).Value)
		
		' Get text field value that should have changed in Supplier: Number
		Dim objOracleTextFieldNumber: Set objOracleTextFieldNumber = GetOracleTextField(objFindInvoicesFormWindow, "SupplierNumber")
		Dim supplierNumberValue: supplierNumberValue = GetOracleTextFieldPropertyValue(objOracleTextFieldNumber,"value")
		Call AssertActualEQUALToExpected(supplierNumberValue, recordSet.Fields(1).Value)

		' Click find button
		Dim objFindOracleButton: Set objFindOracleButton = GetOracleButton(objFindInvoicesFormWindow, "Find")
		Call ClickOracleButton (objFindOracleButton)
		
		' We expect a refreshed Invoice Workbench \(SUPERUSER\) screen here, there should be an assert statement here.
		' Get Form Window :(NOTE): Brackets needs to be escaped in the names (I can put this check within functions)
		Dim objInvoiceWorkbenchWindow: Set objInvoiceWorkbenchWindow = GetOracleFormWindow("Invoice Workbench \(SUPERUSER\)")
		Dim objOracleTable: Set objOracleTable = GetOracleTable(objInvoiceWorkbenchWindow, 70) 'Have 70 columns in this table
		Dim supplierName: supplierName = GetFieldValueFromOracleTable(objOracleTable, 1, 3) 'Row 1, column nr 3 has supplier name
		Call AssertActualEQUALToExpected(supplierName, recordSet.Fields(0).Value)

		' Close the workbench Forms window
		Call CloseOracleFormWindow(objInvoiceWorkbenchWindow)

		recordSet.MoveNext
	Loop

	recordSet.close
	Set recordSet = Nothing
		
End Function

Function CloseInvoiceWorkbenchDemo()   

	Dim objNavigatorWindow: Set objNavigatorWindow = GetOracleFormWindow("Navigator")
	Call CloseOracleFormWindow(objNavigatorWindow)

	Dim objCautionOracleNotification: Set objCautionOracleNotification = GetOracleNotification("Caution")
	Call ApproveOracleNotification(objCautionOracleNotification)
	
End Function
