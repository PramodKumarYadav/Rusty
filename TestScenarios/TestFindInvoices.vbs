
'To avoid errors due to typos in variable names
Option Explicit

'Set up
Function SetUp()   
	
	'Get test environment configuration
	Dim objXMLTestEnv: Set objXMLTestEnv = GetTestEnvConfigurationObject()

	'Close test browsers
	Call CloseTestBrowsers(objXMLTestEnv)

	'Login to Test environment
	Call LoginTestEnvironment(objXMLTestEnv)

	'Open Invoice batches
	Call NavigateToInvoiceBatches()  
	
End Function

'TestCreateInvoices
Function TestCreateInvoices()   
	
	'Create multiple invoices based on test data from create-invoices.csv file
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim fileName: fileName = "InvoiceNrs.csv"	
	Dim testDataRecordSet: testDataRecordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", fileName,"*") 
	Call CreateInvoices(testDataRecordSet)
	
End Function

'TestFindInvoices
Function TestFindInvoices()   
	
	'Find invoices based on test data from find-invoices.csv file
	Call FindInvoice()
	
End Function

'TestFindInvoices
Function TestFindInvoicesToDO()   
	
	'Find invoices based on test data from find-invoices.csv file
	'Todo: This is the way to go
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim fileName: fileName = "FindInvoice.csv"	
	Dim testDataRecordSet: testDataRecordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", fileName,"*") 
	Call FindInvoice(testDataRecordSet)
	
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