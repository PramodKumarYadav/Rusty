' Functions to assert actual vs expected string values
' Usage: Call AssertActualEQUALToExpected(supplierName, recordSet.Fields(0).Value)

Function AssertActualEQUALToExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) = strExpectedValue) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue				 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End If
End Function

Function AssertActualNOTEQUALToExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) <> strExpectedValue) Then
		Reporter.ReportEvent micPass,"Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue				 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualGREATERThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) >strExpectedValue) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue				 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualLESSERThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) <strExpectedValue) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue			 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualGREATEROREQUALThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) >= strExpectedValue) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualLESSEROREQUALThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) <= strExpectedValue) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertREGEXP(strActualValue,strExpectedValue)
	
	Dim regEx : Set regEx = New RegExp
	regEx.Pattern = strExpectedValue
	Dim objFlag: Set objFlag = regEx.Execute(strActualValue)  
	If objFlag.count >0 Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	Else 
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End If
	Set objFlag=Nothing
	Set regEx=Nothing
	
End Function