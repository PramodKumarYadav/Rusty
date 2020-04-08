 'To avoid errors due to typos in variable names
 Option Explicit

 ' Get the XML node object for the environment under test.
 Function GetTestEnvConfigurationObject()
	
	' Create XML object 	
	Dim objXML: Set objXML = CreateObject("MSXML2.DOMDocument.6.0")

	' Get config file path and load XML document
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim pathConfigXML: pathConfigXML = rootDir & "\test-env-config.xml"
	objXML.load(pathConfigXML) 
	
	objXML.async = False
	objXML.validateOnParse = True
	objXML.resolveExternals = False

	' Get Test environment and its correspondig settings node
	Dim testEnv: testEnv = GetSystemEnvironmentVariable("RUSTY_TEST_ENV")
	Dim xPath: xPath = "/Env/"&testEnv
	Dim envNode: Set envNode = objXML.SelectNodes(xPath)
	
	' Return xml object (With Set command)
	Set GetTestEnvConfigurationObject = envNode.item(0)

End Function

' Fetch any key value from the test environemnt.
Function GetXMLChildNodeValue(parentXMLNode, Key)

	Dim xpath: xpath = "./"&Key
	Dim childXMLNode: Set childXMLNode = parentXMLNode.SelectNodes(xpath)
	
	'Return string value (Without Set command)
	GetXMLChildNodeValue = childXMLNode.item(0).nodeTypedValue

End Function

' GetSystem Environment Variables value
Function GetSystemEnvironmentVariable(variableName)
	
	Dim wshShell: Set wshShell = CreateObject( "WScript.Shell" )
	Dim variableValue: variableValue =  wshShell.ExpandEnvironmentStrings("%"&variableName&"%")
	Set wshShell = Nothing

	GetSystemEnvironmentVariable = variableValue

End Function