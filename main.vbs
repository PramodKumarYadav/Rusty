'To avoid errors due to typos in variable names
Option Explicit

'Test Runner: main Entrypoint
Function main()  

	' Set root directory of project (one time action) - 
	' Check if this variable doesnt exist, then set it, else skip. 
	' To-Do: a powershell script to run from root (Till then, use this)

	'Pick tests from "select-tests-to-run.csv" sheet
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")	
	Dim rsSuites: Set rsSuites =  GetCSVFileAsRecordSet(rootDir, "select-test-suits-to-run.csv","*") 

	'Pick test suits that are selected to be run in "select-test-suits-to-run.csv" sheet
	Do Until rsSuites.EOF		 	
		Dim selection: selection = rsSuites.Fields("Select").Value
		Dim testSuite: testSuite = rsSuites.Fields("TestSuite").Value	

		' If Test suite is selected to run. Then Run the test scenarios in it 
		If selection = "Yes" Then						
			Dim rsTestScenarios: Set rsTestScenarios =  GetCSVFileAsRecordSet(rootDir & "\TestSuits", testSuite ,"*") 
			Do Until rsTestScenarios.EOF		 
				Dim selectionTS: selectionTS = rsTestScenarios.Fields("Select").Value	
				Dim testScenario: testScenario = rsTestScenarios.Fields("TSName").Value

				' If scenario is selected to run. Then Run it.
				If selectionTS = "Yes" Then				
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

End Function
'Some references:
	'Check here for all the options that you can use with recordset
	'https://www.w3schools.com/asp/ado_ref_recordset.asp

	' Note: if you want to pass parameters in execute statement; use this format (say testEnv)
	' Dim fn: fn = testScenario & " " & """" & testEnv & """" 
	' execute fn
