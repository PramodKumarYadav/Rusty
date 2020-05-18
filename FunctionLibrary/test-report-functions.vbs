' Assumption while running this function is that the full report is stored in UFT environment variables: named as "ReportRecord0,ReportRecord1,ReportRecord2,3 etc..."
' Note, we are using UFT env variable, because it is created/destroyed in run time. We dont have to store/version control it in GIT.
' This variable(s) are one dimensional array variables, where each item is something that want to use for reporting (say testSuite, TestScenario, etc...) 
' At this moment, we are using TestSuite, TestScenario, TestStep,Expected,Actual, Status to report.
' However if you want to scale and add more items, its possible. Just think about doing it consistentatly for all tests. Above items can be consistent for all tests.
Function CreateCSVReport()
     
    ' Create a file to store CSV report
    Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")	
    Set objLogFile = CreateFile(rootDir & "\TestReport\test-report.csv", True)
    
    Dim colNr: colNr = 1
    'TestSuite, TestScenario, TestStep,Expected,Actual, Status (Scalable, if you want to add some more layers, such as Testcase after TestScenario)
    Dim totalColumnsInReport: totalColumnsInReport = 6 
    ' We are using Environment variables of UFT in report, since we 
    For i=0 to Environment("RecordCount")
        record = "ReportRecord" & i 
        
        For each item in Environment(record)
           ' Check the position of item
            divisor = colNr Mod totalColumnsInReport
            
            'If not the last item in the record
            If divisor <> 0 Then
                'Write the record with a comma in the end
                objLogFile.Write item & ","
            End If
            
            'If this is the last item in the record then dont add comma and add end of line instead
            If divisor = 0 Then
                objLogFile.Write item
                objLogFile.Writeline
            End If
            
            colNr = colNr +1
        Next
    Next
    
    ' Close the file
    objLogFile.Close

End Function

Function SetHeaderRecordForTestReport()
    
    ' Add a Header record into UFT properties variable
    Dim header: header = Array("testSuite","testScenario","testStep","expected","actual","status")
    Environment.Value("ReportRecord0") = header

    ' Initialise the record count to 0
    Environment.Value("RecordCount") = 0

End Function

Function SetResultRecordForTestReport(testStep, expected, actual, result)
    
    ' Get the current record count 
    Dim recordCount: recordCount = Environment("RecordCount")

    ' Get the testScuite name
    Dim testSuite: testSuite = Environment("testSuite")

    ' Get the testScenario name
    Dim testScenario: testScenario = Environment("testScenario")

    ' Increment the count by one 
    recordCount = recordCount + 1 

    ' Add a result record into UFT properties variable
    Dim resultRecord: resultRecord = Array(testSuite,testScenario,testStep, expected, actual, result)

    Dim currentRecordName: currentRecordName = "ReportRecord" & recordCount
    Environment.Value(currentRecordName) = resultRecord

    ' Set the record count to this new count
    Environment.Value("RecordCount") = recordCount

    ' For redundancy in reporting, also publish this result to UFT reports (good way to cross check)
    If result = "Pass" Then
        Reporter.ReportEvent micPass, "Check Passed", "Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue
    Else 
        Reporter.ReportEvent micFail,"Check Failed","Expected: " & strExpectedValue & VbCrLf & "Actual: " & strActualValue	 
    End If   	

End Function

Function CreateTestReports()
    
    ' Create a CSV test report at the end
    Call CreateCSVReport()

    ' Create test reports in various formats
    Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")	
	RunPowershellScript(rootDir & "PSScripts\create-test-report-formats.ps1")

End Function
