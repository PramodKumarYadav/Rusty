
'To avoid errors due to typos in variable names
Option Explicit

Function DemoSwitchResponsibilityToItalySUPERUSER()   
	
	Call SwitchResponsibility("Navigator - System Administrator", "IT SUPERUSER")  
	
End Function

' Ex: 'OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function DemoList").Activate "+  Navigate"
Function DemoNavigateToInvoiceBatches()   

	' 'OracleNotification("Error").Approve
		' Dim objOracleNotification: Set objOracleNotification = GetOracleNotification("Error")
		' Call ApproveOracleNotification(objOracleNotification)

	'Navigate to the Invoices form
	'OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Activate "+  Navigate"
	Dim objOracleFormWindow: Set objOracleFormWindow = GetOracleFormWindow("Navigator - IT SUPERUSER")
		Dim objOracleTabbedRegion: Set objOracleTabbedRegion = GetOracleTabbedRegion(objOracleFormWindow, "Functions")
			Dim objGetOracleList: Set objGetOracleList = GetOracleList(objOracleTabbedRegion, "Functions List")
			Call ActivateOracleListItem (objGetOracleList, "+  Navigate")
			Call ActivateOracleListItem (objGetOracleList, "   +  Invoices")
			Call ActivateOracleListItem (objGetOracleList, "      +  DoInquiry")
			Call ActivateOracleListItem (objGetOracleList, "             MyInvoices")
	
End Function

'TestFindInvoices
Function DemoTestFindInvoices()   
	
	'Find invoices based on test data from find-invoices.csv file
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim recordSet: Set recordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", "FindInvoices.csv","*")
	Call FindInvoice(recordSet)
	
End Function

'Find Invoices	
Function DemoFindInvoice(recordSet)   

		Do Until recordSet.EOF		
			
			' Get Form Window
			Dim objFindInvoicesFormWindow: Set objFindInvoicesFormWindow = GetOracleFormWindow("Find Invoices")
			
			' Set text field with Trading Partner: Name
			Dim objOracleTextFieldName: Set objOracleTextFieldName = GetOracleTextField(objFindInvoicesFormWindow, "Invoice: Name")
			Call SetOracleTextFieldValue(objOracleTextFieldName, recordSet.Fields(0).Value)
			
			' Get text field value that should have changed in Supplier: Number
			Dim objOracleTextFieldNumber: Set objOracleTextFieldNumber = GetOracleTextField(objFindInvoicesFormWindow, "Supplier: Number")
			Dim supplierNumberValue: supplierNumberValue = GetOracleTextFieldPropertyValue(objOracleTextFieldNumber,"value")
			Call AssertActualEQUALToExpected(supplierNumberValue, recordSet.Fields(1).Value)

			' Click find button
			Dim objFindOracleButton: Set objFindOracleButton = GetOracleButton(objFindInvoicesFormWindow, "Find")
			Call ClickOracleButton (objFindOracleButton)
			
			' We expect a refreshed Invoice Workbench \(IT SUPERUSER\) screen here, there should be an assert statement here.
			' Get Form Window :(NOTE): Brackets needs to be escaped in the names (I can put this check within functions)
			Dim objInvoiceWorkbenchWindow: Set objInvoiceWorkbenchWindow = GetOracleFormWindow("Invoice Workbench \(IT SUPERUSER\)")
			Dim objOracleTable: Set objOracleTable = GetOracleTable(objInvoiceWorkbenchWindow, 58) 'Have 58 columns in this table
			Dim supplierName: supplierName = GetFieldValueFromOracleTable(objOracleTable, 1, 6) 'Row 1, column nr 6 has supplier name
			Call AssertActualEQUALToExpected(supplierName, recordSet.Fields(0).Value)

			' Note: You dont have to close the workbench Forms window. The find invoices window is just behind this workbench window (alive). 
			' You can search the next invoice directly.

			recordSet.MoveNext
		Loop

		recordSet.close
		Set recordSet = Nothing
		
End Function

Function DemoExitOracleApplicationFromItalySUPERUSER()   

	Dim objNavigatorWindow: Set objNavigatorWindow = GetOracleFormWindow("Navigator - IT SUPERUSER")
	Call SelectMenuOracleFormWindow(objNavigatorWindow, "File->Exit Oracle Applications")

	Dim objCautionOracleNotification: Set objCautionOracleNotification = GetOracleNotification("Caution")
	Call ApproveOracleNotification(objCautionOracleNotification)
	
End Function

' Below you see various ways of calling test data

'TestFindInvoices using a csv file
Function DemoTestFindInvoicesViaCSV()   
	
	'Find invoices based on test data from find-invoices.csv file
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim recordSet: Set recordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", "FindInvoices.csv","*")
	Call FindInvoice(recordSet)
	
End Function

'TestFindInvoices using values from database table
Function DemoTestFindInvoicesViaDB()   
	
	'Find invoices based on test data from database table
	Dim sql: sql = "select * from ab.cd_invoices FETCH FIRST 2 ROWS ONLY"
	Dim recordSet: Set recordSet = GetDBTableAsRecordSet(sql)
	Call FindInvoice(recordSet)
	
End Function

'TestCreateInvoices
Function DemoTestCreateInvoices()   
	
	'Create multiple invoices based on test data from create-invoices.csv file
	' If in future, we have subdirectories to look for test data, we may want to give path of directory (thats why we pass it as an argument)
	' If all files live in rootDir\TestData than this would not be needed to give as a parameter and can be known within the function.
	' I expect that in future, we may have sub directories for different type of tests. 
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim recordSet: Set recordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", "InvoiceNrs.csv","*")
	Call CreateInvoices(testDataRecordSet)
	
End Function
