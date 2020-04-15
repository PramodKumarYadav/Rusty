
'To avoid errors due to typos in variable names
Option Explicit

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

	'Open Invoice batches
	Call NavigateToInvoiceBatches()  
	
End Function

'TestCreateInvoices
Function TestCreateInvoices()   
	
	'Create multiple invoices based on test data from create-invoices.csv file
	' If in future, we have subdirectories to look for test data, we may want to give path of directory (thats why we pass it as an argument)
	' If all files live in rootDir\TestData than this would not be needed to give as a parameter and can be known within the function.
	' I expect that in future, we may have sub directories for different type of tests. 
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim recordSet: Set recordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", "InvoiceNrs.csv","*")
	Call CreateInvoices(testDataRecordSet)
	
End Function

'TestFindInvoices
Function TestFindInvoices()   
	
	' Get a few records from the table for which we expect to find them from oracle forms
	' Dim sql: sql = "select * from gl.gl_ledgers FETCH FIRST 2 ROWS ONLY"
	' Dim recordSet: Set recordSet = GetDBTableAsRecordSet(sql)
	' do until recordSet.EOF
		
	' 	msgbox recordSet.Fields(0).Value
	' 	msgbox recordSet.Fields(1).Value
	' 	msgbox recordSet.Fields(2).Value
		
	' 	recordSet.MoveNext
	' loop

	' recordSet.close
	' Set recordSet = Nothing

	'Find invoices based on test data from find-invoices.csv file
	Call FindInvoiceOld()
	
End Function

'TestFindInvoicesViaCSV
Function TestFindInvoicesViaCSV()   
	
	'Find invoices based on test data from find-invoices.csv file
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim recordSet: Set recordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", "FindInvoices.csv","*")
	Call FindInvoice(recordSet)
	
End Function

'TestFindInvoicesViaDB
Function TestFindInvoicesViaDB()   
	
	'Find invoices based on test data from database table
	Dim sql: sql = "select * from gl.gl_ledgers FETCH FIRST 2 ROWS ONLY"
	Dim recordSet: Set recordSet = GetDBTableAsRecordSet(sql)
	Call FindInvoice(recordSet)
	
End Function

'Tear Down
Function TearDown()   
	
	'Get test environment configuration
	Dim objXMLTestEnv: Set objXMLTestEnv = GetTestEnvConfigurationObject()

	'Logout from browser
	Call LogoutBrowser()
	
	'Close test browsers
	Call CloseTestBrowsers(objXMLTestEnv)
	
End Function
