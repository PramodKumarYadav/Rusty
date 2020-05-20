' Copy this script in a UFT - Test - Action
' This script will load all libraries and call main test
' Note: If you want to debug in UFT, loading libraries like this doesnt work properly with debugging. 
' In that case, manually add libraries in a test action with only one line -> call to main()

'Get root directory of your framework - Rusty
Dim rootDir: rootDir =  CreateObject( "WScript.Shell" ).ExpandEnvironmentStrings("%RUSTY_HOME%")

' Load UFT functions library (this contains fn to load all other libraries)
LoadFunctionLibrary rootDir & "\uft-functions.vbs"

' Put them in the order they show in vscode (this way it would be easy to find a missing dir)
Call LoadAllFunctionLibraries(rootDir & "\FunctionLibrary")
Call LoadAllFunctionLibraries(rootDir & "\TestScenarios")

' Load main entrypoint script
LoadFunctionLibrary rootDir & "\main.vbs"

' This wilL now run all selected tests in "select-test-suits-to-run.csv"
Call main() 