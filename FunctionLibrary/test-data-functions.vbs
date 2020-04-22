' To make this work, download a 32 bit driver for your oracle database. Detailed instructions are in Readme.md file.
' Tested OKay from both UFT and running vbs script from 32 bit command prompt
' Usage example: 
	' Dim sql: sql = "select * from employees.table_name FETCH FIRST 2 ROWS ONLY"
	' Dim recordSet: Set recordSet = GetDBTableAsRecordSet(sql)
	' do until recordSet.EOF
		
	' 	msgbox recordSet.Fields(0).Value
	' 	msgbox recordSet.Fields(1).Value

	' 	recordSet.MoveNext
	' loop

	' recordSet.close
	' Set recordSet = Nothing
	Function GetDBTableAsRecordSet(sql)

	' Get connection string
	Dim connString: connString = GetConnectionString()

	' Create and open the connection 
	Dim connection: Set connection  = SetConnection(connString)

	' Get recordset 
	Set GetDBTableAsRecordSet = GetRecordSet(connection, sql)

	' Close the connection
	CloseConnection(connection)

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

Function SetConnection(connString)

	' Create a connection and open connection 
	Dim connection: Set connection  = CreateObject("ADODB.Connection")
	connection.Open connString

	Set SetConnection = connection

End function

Function GetRecordSet(connection, sql)

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

	'Return the detached recordSet (Recordset will be closed in the calling/using function.)
	Set GetRecordSet = recordSet

End function

Function CloseConnection(connection)

	' Close the connection
	connection.close
	Set connection = Nothing

End function

'Usage:
' Get root directory for test data
' Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
' Dim pathTestDataDir: pathTestDataDir = rootDir & "\TestData"
' ex:1 (to fetch all records): Call GetCSVFileAsRecordSet(pathTestDataDir, fileName,"*") 
'Usage ex:2 (to fetch limited records): Call GetCSVFileAsRecordSet(pathTestDataDir, fileName,2) 
Function GetCSVFileAsRecordSet(pathTestDataDir, fileName, iterations)  

	' Get connection string
	Dim connString: connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&pathTestDataDir&";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"

	' Create and open the connection 
	Dim connection: Set connection  = SetConnection(connString)

	' Prepare SQL based on the user request 
	Dim sql
	If iterations = "*" then
		sql = "SELECT * FROM ["&fileName&"]"
	Else
		sql = "SELECT TOP "&iterations&" * FROM ["&fileName&"]"
	End If

	' Get recordset 
	Set GetCSVFileAsRecordSet = GetRecordSet(connection, sql)

	' Close the connection
	CloseConnection(connection)

End function