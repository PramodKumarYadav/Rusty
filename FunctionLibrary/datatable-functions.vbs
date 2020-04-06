Function GetParameterValueFromDataSheet(strParameterName, strSheetType)
	
	'Return the value of parameter based on global or local
	If strSheetType = "Local" Then		
		GetParameterValueFromDataSheet = DataTable.Value(strParameterName, dtLocalSheet)
	Else
		GetParameterValueFromDataSheet = DataTable(strParameterName)  
	End If	
	
End Function