'To avoid errors due to typos in variable names
Option Explicit

' Set root directory of project (Need to run only one time : first time as part of your setup before you start running the tests.)
Function setRootDirectory(yourProjectDirectoryPath)  

	RunPowershellScript(yourProjectDirectoryPath & "\set-project-root.ps1")

End Function

'Test Runner: main Entrypoint
Function main()  

	' To capture any unhandled/unexpected error while running the tests.
	' On Error Resume Next

	' Start reporting
	SetHeaderRecordForTestReport()

	'Pick tests from "select-tests-to-run.csv" sheet
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")	
	Dim sql: sql = "SELECT * FROM [select-test-suits-to-run.csv]"
	Dim rsSuites: Set rsSuites = GetCSVFileAsRecordSet(rootDir, sql)

	'Pick test suits that are selected to be run in "select-test-suits-to-run.csv" sheet
	Do Until rsSuites.EOF		 	
		Dim selection: selection = rsSuites.Fields("Select").Value
		Dim testSuite: testSuite = rsSuites.Fields("TestSuite").Value	
		
		' If Test suite is selected to run. Then Run the test scenarios in it 
		If selection = "Yes" Then	
			
			' Set UFT parameter to report on this testSuite
			Environment("testSuite")= testSuite

			Dim sqlTS: sqlTS = "SELECT * FROM [" & testSuite & "]"
			Dim rsTestScenarios: Set rsTestScenarios = GetCSVFileAsRecordSet(rootDir & "\TestSuits", sqlTS)
			Do Until rsTestScenarios.EOF		 
				Dim selectionTS: selectionTS = rsTestScenarios.Fields("Select").Value	
				Dim testScenario: testScenario = rsTestScenarios.Fields("TSName").Value

				' If scenario is selected to run. Then Run it.
				If selectionTS = "Yes" Then	
					
					' Set UFT parameter to report on this testScenario
					Environment("testScenario")= testScenario

					execute testScenario
				End If

				' Once done, go to next record (Test scenario)
				rsTestScenarios.MoveNext
			Loop			
			' At the end close recordset and release the object. 
			rsTestScenarios.close
			Set rsTestScenarios = Nothing
		End If

		' Once done, go to the next record (Next TestSuite)
		rsSuites.MoveNext
	Loop

	' At the end close recordset and release the object. 
	rsSuites.close
	Set rsSuites = Nothing
	
	' In the end, Create test reports in differnt formats 
	Call CreateTestReports()

	' If Err.Number <> 0 Then
	' 	Reporter.ReportEvent micFail,"There was an unexpected/unhandled error encountered during the run. Err.Description is: [" & Err.Description & "]", "debug and fix to handle this better next time!"
	' 	Err.Clear
	' 	ExitTest
    ' End If

End Function
'Some references:
	'Check here for all the options that you can use with recordset
	'https://www.w3schools.com/asp/ado_ref_recordset.asp

	' Note: if you want to pass parameters in execute statement; use this format (say testEnv)
	' Dim fn: fn = testScenario & " " & """" & testEnv & """" 
	' execute fn
