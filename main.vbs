'To avoid errors due to typos in variable names
Option Explicit

'Test Runner: main Entrypoint
Function main()  

	'Choose a test environment 
	Dim testEnv: testEnv = "Dev"

	'Pick tests that are selected to be run in "select-tests-to-run.csv" sheet
	'Todo: Make this path user specific.	
	 Dim pathParentDir: pathParentDir = "C:\Users\nlpyad1\UFT\Rusty"
	 Dim fileName: fileName = "select-tests-to-run.csv"
	 
	 Dim recordSetTS: Set recordSetTS =  GetTestData(pathParentDir, fileName,"*")   
	 
	'Check here for all the options that you can use with recordset
	'https://www.w3schools.com/asp/ado_ref_recordset.asp
	 Do Until recordSetTS.EOF

		Dim testScenario: testScenario = recordSetTS.Fields(0).Value
		Dim selection: selection = recordSetTS.Fields(1).Value

		' If scenario is selected to run. Then Run it passing the testEnv value as a parameter.
		If selection = "Yes" Then
			Dim fn: fn = testScenario & " " & """" & testEnv & """"
			execute fn
		End If
		
		' Once done, go to next record (Test scenario)
		recordSetTS.MoveNext
	Loop
	
	'At the end close recordset and release the object. 
	recordSetTS.close
	Set recordSetTS = Nothing

End Function