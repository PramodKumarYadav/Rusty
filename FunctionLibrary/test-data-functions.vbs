'To avoid errors due to typos in variable names
Option Explicit

'Usage ex:1 (to fetch all records): Call GetCSVFileAsRecordSet(pathParentDir, fileName,"*") 
'Usage ex:2 (to fetch limited records): Call GetCSVFileAsRecordSet(pathParentDir, fileName,2) 
Function GetCSVFileAsRecordSet(pathTestDataDir, fileName, iterations)   
	
	Dim connection: Set connection = CreateObject("ADODB.Connection")
	connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&pathTestDataDir&";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
	connection.open
	
	'Specify how many records to pick from this test data file.
	Dim sql
	If iterations = "*" then
		sql = "SELECT * FROM ["&fileName&"]"
	Else
		sql = "SELECT TOP "&iterations&" * FROM ["&fileName&"]"
	End If
	
	Dim recordSet: Set recordSet = CreateObject("ADODB.Recordset")
	Set recordSet.ActiveConnection = connection

	Const adOpenStatic = 3
	Const adLockOptimistic = 3
	Const adUseClient = 3
	recordSet.CursorLocation = adUseClient
	recordSet.CursorType = adOpenStatic
	recordSet.LockType = adLockOptimistic

	' Run the query.
	recordSet.Source = sql
	recordSet.Open

	' Disconnect the recordset.
	Set recordSet.ActiveConnection = Nothing

	'Return the detached recordSet (Recordset will be closed in the calling function.)
	Set GetCSVFileAsRecordSet = recordSet

	connection.close
	
End Function

Function GetCSVFileAsRecordSetBackUp(pathTestDataDir, fileName, iterations)   
	
	Dim ado: Set ado = CreateObject("ADODB.Connection")
	ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&pathTestDataDir&";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
	ado.open
	
	'Specify how many records to pick from this test data file.
	Dim sql
	If iterations = "*" then
		sql = "SELECT * FROM [InvoiceNrs.csv]"
	Else
		sql = "SELECT TOP "&iterations&" * FROM [InvoiceNrs.csv]"
	End If
	
	Dim recordSet: Set recordSet = ado.Execute(sql)
	
	'I can now return this recordset object for tests to use it.
	
	'Return the recordSet
	Set GetCSVFileAsRecordSet = recordSet
	
	'In the test, we can iterate over whatever iteration was selected, to do the work we want to do (as shown below)
	'for now kept here for reference, until I start making use of this.
	Dim field1, field2
	Do Until recordSet.EOF

		field1 = recordSet.Fields(0).Value
		field2 = recordSet.Fields(1).Value

		' Use your fields here in test.
		
		' Once done, go to next record
		recordSet.MoveNext
	Loop

	recordSet.close
	ado.close
	
End Function

