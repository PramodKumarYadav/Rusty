'Todo: Still need to make this work with oracle, but this should work for all practical purposes.
Function GetDBTableAsRecordSet(connString, sql)

	Dim connection: Set connection  = CreateObject("ADODB.Connection")
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

End function