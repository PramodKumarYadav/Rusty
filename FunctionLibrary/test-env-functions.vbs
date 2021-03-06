 'To avoid errors due to typos in variable names
 Option Explicit

 ' Get the parent XML object that contains the details you are interested in
 Function GetXMLNodeObject(xmlName, xPath)
	
	' Create XML object 	
	Dim objXML: Set objXML = CreateObject("MSXML2.DOMDocument.6.0")

	' Get config file path and load XML document
	Dim rootDir: rootDir = GetSystemEnvironmentVariable("RUSTY_HOME")
	Dim pathConfigXML: pathConfigXML = rootDir & "\" & xmlName
	objXML.load(pathConfigXML) 
	
	objXML.async = False
	objXML.validateOnParse = True
	objXML.resolveExternals = False

	' Get Test environment and its correspondig settings node
	Dim envNode: Set envNode = objXML.SelectNodes(xPath)
	
	' Return xml object (With Set command)
	Set GetXMLNodeObject = envNode.item(0)

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

 ' Get the XML node object for the environment under test.
 Function GetTestEnvConfigurationObject()
	
	' Get chosen Test environment value to run tests
	Dim chosenTestEnv: chosenTestEnv = GetChosenTestEnvValue()
	
	' Get chosenTestEnv XML object for this chosen test env
	Set GetTestEnvConfigurationObject = GetXMLNodeObject("config-test-env.xml", "/Env/"&chosenTestEnv)

End Function

 ' Get the XML node object for the user secrets.
 Function GetChosenTestEnvValue()
	
	' Get Test environment to run tests
	Dim xmlName: xmlName = "config-user-secrets.xml"
	Dim xPath: xPath = "/Env"
	Dim objXMLUser: Set objXMLUser = GetXMLNodeObject(xmlName, xPath)
	
	' Since its a value, no use of Set function
	GetChosenTestEnvValue = GetXMLChildNodeValue(objXMLUser, "ChosenTestEnv")

End Function

 ' Get the XML node object for the user secrets.
 ' Reason of having two seperate config files for test env and user is since different users will have test env config same but user config different. 
 ' We dont want to accidentally push user secrets in GIT. We can also git ignore this file to be avoid pushing secrets on git.
 Function GetUserConfigurationObject()
	
	' Get chosen Test environment value to run tests
	Dim chosenTestEnv: chosenTestEnv = GetChosenTestEnvValue()
	
	' Get chosenUser XML object for this chosen test env
	Set GetUserConfigurationObject = GetXMLNodeObject("config-user-secrets.xml", "/Env/"&chosenTestEnv)

End Function


