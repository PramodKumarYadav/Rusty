
'To avoid errors due to typos in variable names
Option Explicit

'TestFindInvoices
Function TestFindInvoices()   
	
	'Find invoices based on test data from find-invoices.csv file
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim recordSet: Set recordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", "FindInvoices.csv","*")
	Call FindInvoice(recordSet)
	
End Function

'TestFindInvoicesViaDB
Function TestFindInvoicesViaDB()   
	
	'Find invoices based on test data from database table
	Dim sql: sql = "select * from ab.cd_invoices FETCH FIRST 2 ROWS ONLY"
	Dim recordSet: Set recordSet = GetDBTableAsRecordSet(sql)
	Call FindInvoice(recordSet)
	
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