' Tested OKay from my local personal machine. Yet to make it work from my work machine. 
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

' This worked from my local machine. And it seems OLEDB is the way to go when  using VB to connect. See this and look for section (Oracle XE, VB6 ADO):
' https://www.connectionstrings.com/oracle-provider-for-ole-db-oraoledb/

connString = "Provider=OraOLEDB.Oracle;dbq=localhost:1521/ORCL;User ID=system;Password=Securid30;"
' Below sql is generic and gives oracle product details. Should run on your machine as well unchanged. 
' If you are not sure of a database name, use this to test connection first.
sql = "SELECT * FROM PRODUCT_COMPONENT_VERSION"

Set recordset =  GetDBTableAsRecordSet(connString, sql)
do until recordSet.EOF
	
	msgbox recordSet.Fields(0).Value
	' msgbox recordSet.Fields(1).Value
	' msgbox recordSet.Fields(2).Value
	
	recordSet.MoveNext
loop

recordSet.close
Set recordSet = Nothing
	