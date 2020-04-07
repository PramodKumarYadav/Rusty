'To avoid errors due to typos in variable names
Option Explicit

'Usage ex:1 (to fetch all records): Call GetTestData(pathParentDir, fileName,"*") 
'Usage ex:2 (to fetch limited records): Call GetTestData(pathParentDir, fileName,2) 
Function GetTestData(pathTestDataDir, fileName, iterations)   
	
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
	Set GetTestData = recordSet
	
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


