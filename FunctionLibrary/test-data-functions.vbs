' To make this work, download a 32 bit driver for your oracle database. Detailed instructions are in Readme.md file.
' Tested OKay from both UFT and running vbs script from 32 bit command prompt
' Usage example: 
	' Dim sql: sql = "select * from employees.table_name FETCH FIRST 2 ROWS ONLY"
	' Dim recordSet: Set recordSet = GetDBTableAsRecordSet(sql)
	' do until recordSet.EOF
		
	' 	msgbox recordSet.Fields("EmpName").Value
	' 	msgbox recordSet.Fields("EmpNr").Value

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

	Dim recordSet: Set recordSet = SetRecordSet(connection)

	' Run the query.
	recordSet.Source = sql
	recordSet.Open

	' Disconnect the recordset.
	Set recordSet.ActiveConnection = Nothing

	'Return the detached recordSet (Recordset will be closed in the calling/using function.)
	Set GetRecordSet = recordSet

End function

Function SetRecordSet(connection)

	Dim recordSet: Set recordSet = CreateObject("ADODB.Recordset")
	Set recordSet.ActiveConnection = connection

	' Const adOpenStatic = 3
	' Const adLockOptimistic = 3
	' Const adUseClient = 3
	recordSet.CursorLocation = 3 	' adUseClient
	recordSet.CursorType = 3		' adOpenStatic
	recordSet.LockType = 3 			' adLockOptimistic

	Set SetRecordSet = recordSet

End function

Function CloseConnection(connection)

	' Close the connection
	connection.close
	Set connection = Nothing

End function

' Example to set SQL and call this function.
' Ex 1: Dim sql: sql = "SELECT TOP 2 * FROM  [find-invoices.csv]"
' Ex 2: Dim sql: sql = "SELECT * FROM [find-invoices.csv] where SupplierNumber = '9998823977'"
' Dim recordSet: Set recordSet = GetCSVFileAsRecordSet(rootDir & "\TestData", sql)
Function GetCSVFileAsRecordSet(pathTestDataDir, sql)  

	' Get connection string
	Dim connString: connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&pathTestDataDir&";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"

	' Create and open the connection 
	Dim connection: Set connection  = SetConnection(connString)

	' Get recordset 
	Set GetCSVFileAsRecordSet = GetRecordSet(connection, sql)

	' Close the connection
	CloseConnection(connection)

End function

' Design decision: This is from a usage perspective opposite of Function GetDBTableAsRecordSet(sql)
' In Function GetDBTableAsRecordSet(sql), we give a sql statement and we expect zero, one or more records to get back and work with.
' Normally when we want to execute statements (say insert, update or delete), we often want to do it for one or more records and expect no records back.
' Passing sqls as recordset will enable you to handle more than one sql inserts in one go. 
' You pick which sqls to pass to this function and update them all in one single go.

' Note: This function assumes that there is a column name "SQL", in the passed recordSetSQL
Function ExecuteSQLStatements(recordSetSQL)

	' Get connection string
	Dim connString: connString = GetConnectionString()

	' Create and open the connection 
	Dim connection: Set connection  = SetConnection(connString)

	' Create recordset object
	Dim recordSet: Set recordSet = SetRecordSet(connection)

	' Execute each of the SQL statements from recordSetSQL
	Do Until recordSetSQL.EOF	
		Dim sql: sql = recordSetSQL.Fields("SQL").Value
		recordSet.Source = sql
		recordSet.Open

		recordSetSQL.MoveNext
	Loop

	' Commit the changes, otherwise in next run, you will get weired errors.
	recordSet.Source = "commit;"
	recordSet.Open

	' User closes the recordset so close it now.
	recordSetSQL.close
	Set recordSetSQL = Nothing

	' Disconnect the recordset.
	Set recordSet.ActiveConnection = Nothing

	' No need to pass recordset (we will validate seperately if this was succesful or not)
	Set recordSet = Nothing

	' Close the connection
	CloseConnection(connection)

End function



' Example usage: 
' Call ExecuteStoredProcedure("runSampleProgram18", "XXTEST_NCMAISUB.submit_group_import")
Function ExecuteStoredProcedure(tag, storedProcedureName)

	On Error Resume Next

	' Get connection string
	Dim connString: connString = GetConnectionString()

	' Create and open the connection 
	Dim connection: Set connection  = SetConnection(connString)

	' Set the command
	Dim cmd: Set cmd = SetCommand(connection)
	
	' Prepare the stored procedure
	Call PrepareStoredProcedure(cmd, storedProcedureName)

	' Create parameters for this stored procedure
	Call CreateAndAppendParameters(cmd, tag, storedProcedureName)

	' Execute the stored procedure
	Call ExecuteCommand(cmd)

	' Destroy the object
	Set cmd = Nothing

	' Close the connection
	CloseConnection(connection)

	 If Err.Number <> 0 Then
	 	Reporter.ReportEvent micDone,"Due to timeouts,we expect errors in getting proper return codes. However procedure, itself goes okay and hence this step is micDone and not micFail. Here are details of error.Err.Description is: [" & Err.Description & "]", "Check your data and debug"
	 	Err.Clear
	 	ExitTest
     End If

End function

' Set the command
Function SetCommand(connection)

	Dim cmd: Set cmd = CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = connection

	Set SetCommand = cmd

End function

' Prepare the stored procedure
Function PrepareStoredProcedure(cmd, storedProcedureName)

	cmd.CommandType = 4  'adCmdStoredProc
	cmd.CommandText = storedProcedureName ' Stored procedure name (say): "[dbo].[sptestproc]"

End function

' Create and append parameters
Function CreateAndAppendParameters(cmd, tag, storedProcedureName)

	' Get recordset for procedure definition (based on storedProcedureName)
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim sql: sql = "SELECT * FROM [procedure-definition.csv] where ProcedureName = '" & storedProcedureName & "'"
	Dim recordSetSignature: Set recordSetSignature = GetCSVFileAsRecordSet(rootDir & "\TestData\StoredProcedures", sql)
	Dim countSignatureRecords: countSignatureRecords = recordSetSignature.RecordCount

	' Get recordset for procedure values (based on tag)
	sql = "SELECT * FROM [stored-procedures.csv] where Tag = '" & tag & "'"
	Dim recordSetTag: Set recordSetTag = GetCSVFileAsRecordSet(rootDir & "\TestData\StoredProcedures", sql)
	Dim countTagRecords: countTagRecords = recordSetTag.RecordCount

	' Assert that the number of records is same in both (i.e. if the signature matches with instance called)
	' Also if this doesnt match, we would like to abort (deviate from default behaviour of continue)
	' So, get the function return value and decide accordingly
	Dim testStep: testStep = "Assert that nr of records in procedure signature and called instance are same"
	Dim bothEqual: bothEqual =  AssertActualIsEQUALToExpected(testStep, countSignatureRecords,countTagRecords)
	If bothEqual <> True Then
		Call AbortTest()
	End If

	' Iterate for each record and create parameter
	do until recordSetSignature.EOF
			
		' name,type,direction,size,value (Firt 4 comes from procedure-definition.csv)
		Dim name: name = recordSetSignature.Fields("Name").Value
		Dim typeVal: typeVal = recordSetSignature.Fields("Type").Value
		Dim direction: direction = recordSetSignature.Fields("Direction").Value
		Dim size: size = recordSetSignature.Fields("Size").Value
		
		' name,type,direction,size,value (5th (value) come from stored-procedure.csv)
		Dim value: value = recordSetTag.Fields("Value").Value

		' Info: While using values from csv as recordset, the parameters seems to be set correct. 
		' So even if something may be shown as string in debug mode, while parsing from csv as recordset (say a number field)
		' It seems to be inserted correctly when creating parameters in next step. 

		' Create and append parameter
		cmd.Parameters.Append cmd.CreateParameter(name,typeVal,direction,size,value)

		' Move recordsets 
		recordSetSignature.MoveNext
		recordSetTag.MoveNext
	loop

End function

' Execute the stored procedure
Function ExecuteCommand(cmd)

	cmd.Execute

End function
