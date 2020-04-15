' To make this work, download a 32 bit driver for your oracle database. Detailed instructions are in Readme.md file.
' Tested OKay from both UFT and running vbs script from 32 bit command prompt
' Usage example: 
	' Dim sql: sql = "select * from employees.table_name FETCH FIRST 2 ROWS ONLY"
	' Dim recordSet: Set recordSet = GetDBTableAsRecordSet(sql)
	' do until recordSet.EOF
		
	' 	msgbox recordSet.Fields(0).Value
	' 	msgbox recordSet.Fields(1).Value
	' 	msgbox recordSet.Fields(2).Value
		
	' 	recordSet.MoveNext
	' loop

	' recordSet.close
	' Set recordSet = Nothing
Function GetDBTableAsRecordSet(sql)

	Dim connection: Set connection  = CreateObject("ADODB.Connection")

	Dim connString: connString = GetConnectionString()
	connection.Open connString

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
	Set GetDBTableAsRecordSet = recordSet

	connection.close
	Set connection = Nothing

End function

' Get connection string from your config-user-secrets.xml file. A valid connection string looks like below.
' connString = "DRIVER={Oracle in OraClient12Home1_32bit};DBQ=ab01.myCompany.com:1521/ORAB;Trusted_Connection=Yes;UID=your-db-user-id;Password=your-db-password"
Function GetConnectionString()
	
	'Get user configuration 
	Dim objXMLUser: Set objXMLUser = GetUserConfigurationObject()

	'Get values for connection string
	Dim dbDriverAndConn: dbDriverAndConn = GetXMLChildNodeValue(objXMLUser, "dbDriverAndConn")
	Dim dbUser: dbUser = GetXMLChildNodeValue(objXMLUser, "dbUser")
	Dim dbPassword: dbPassword = GetXMLChildNodeValue(objXMLUser, "dbPassword")

	GetConnectionString = dbDriverAndConn & ";UID=" & dbUser & ";Password=" & dbPassword

	Set objXMLUser = Nothing
End Function

'Usage:
' Get root directory for test data
' Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
' Dim pathTestDataDir: pathTestDataDir = rootDir & "\TestData"
' ex:1 (to fetch all records): Call GetCSVFileAsRecordSet(pathTestDataDir, fileName,"*") 
'Usage ex:2 (to fetch limited records): Call GetCSVFileAsRecordSet(pathTestDataDir, fileName,2) 
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
	Set connection = Nothing
End Function