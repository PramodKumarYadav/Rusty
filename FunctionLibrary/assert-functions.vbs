' Functions to assert actual vs expected string values
' NOTE01: Anything worth testing is worth reporting. Or rather anything that is tested must be reported. Due to this pricniple, 
' You will see that after each test, we call the report function, which uses all the parameters from test and one more (result)

' NOTE02: The default behaviour for assert functions is to report an assertion pass/failure and then continue (not abort)
' This makes sense that if there are functional errors and not application errors, we dont want to stop other tests because one test failed.
' However, if in a specific situation, you do want to stop tests if assertions dont match, it is possible. 
' In that case, use the return value of the function and abort if it is not true. Example is given below.
' Dim bothEqual: bothEqual =  AssertActualIsEQUALToExpected(testStep, countSignatureRecords,countTagRecords)
' If bothEqual <> True Then
' 	Call AbortTest()
' End If

' Usage: Call AssertActualIsEQUALToExpected(testStep, supplierName, recordSet.Fields("SupName").Value)
Function AssertActualIsEQUALToExpected(testStep, strActualValue, strExpectedValue)
	Dim result: result = "Fail"	'Default is fail. Pass only when we go in the if loop. All unexpected situations are Fail.
	
	If (Trim(strActualValue) = Trim(strExpectedValue)) Then
		result = "Pass" 
		AssertActualIsEQUALToExpected = True
	End If

	' Report the result
	Call SetResultRecordForTestReport(testStep, strActualValue,strExpectedValue, result)
End Function

Function AbortTest()
	ExitTest
End Function

Function AssertActualIsNOTEQUALToExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) <> Trim(strExpectedValue)) Then
		Reporter.ReportEvent micPass,"Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
		AssertActualIsNOTEQUALToExpected = True 				 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualIsGREATERThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) >Trim(strExpectedValue)) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
		AssertActualIsGREATERThanExpected = True 				 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualIsLESSERThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) <Trim(strExpectedValue)) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
		AssertActualIsLESSERThanExpected = True 			 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualIsGREATEROREQUALThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) >= Trim(strExpectedValue)) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
		AssertActualIsGREATEROREQUALThanExpected = True 
	Else
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End IF
End Function

Function AssertActualIsLESSEROREQUALThanExpected(strActualValue,strExpectedValue)
	If (Trim(strActualValue) <= Trim(strExpectedValue)) Then
		Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
		AssertActualIsLESSEROREQUALThanExpected = True 
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
		AssertREGEXP = True 
	Else 
		Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
	End If
	Set objFlag=Nothing
	Set regEx=Nothing
	
End Function
