'To avoid errors due to typos in variable names
Option Explicit

'Test Runner: main Entrypoint
Function main()  

	' Set root directory of project (one time action) - 
	' Check if this variable doesnt exist, then set it, else skip. 
	' To-Do: a powershell script to run from root (Till then, use this)

	'Pick tests from "select-tests-to-run.csv" sheet
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")	
	Dim recordSetTS: Set recordSetTS =  GetCSVFileAsRecordSet(rootDir, "select-tests-to-run.csv","*") 

	'Pick tests that are selected to be run in "select-tests-to-run.csv" sheet
	Do Until recordSetTS.EOF		 
		Dim testScenario: testScenario = recordSetTS.Fields(0).Value
		Dim selection: selection = recordSetTS.Fields(1).Value

		' If scenario is selected to run. Then Run it.
		If selection = "Yes" Then				
			execute testScenario
		End If

		' Once done, go to next record (Test scenario)
		recordSetTS.MoveNext
	Loop
	
	'At the end close recordset and release the object. 
	recordSetTS.close
	Set recordSetTS = Nothing

End Function
'Some references:
	'Check here for all the options that you can use with recordset
	'https://www.w3schools.com/asp/ado_ref_recordset.asp

	' Note: if you want to pass parameters in execute statement; use this format (say testEnv)
	' Dim fn: fn = testScenario & " " & """" & testEnv & """" 
	' execute fn
