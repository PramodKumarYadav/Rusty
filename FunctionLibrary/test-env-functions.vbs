 'To avoid errors due to typos in variable names
 Option Explicit

 ' Get the XML node object for the environment under test.
 Function GetTestEnvConfigurationObject(pathXML, testEnv)

	'Get xml content from document 
	Dim fileObject: Set fileObject = OpenFile(pathXML, 1, False)
	Dim xmlString: xmlString = fileObject.Readall()	
		
	Dim objXML: Set objXML = CreateObject("MSXML2.DOMDocument.6.0")
	objXML.loadXML(xmlString) 
	
	objXML.async = False
	objXML.validateOnParse = True
	objXML.resolveExternals = False

	Dim xPath: xPath = "/Env/"&testEnv
	Dim envNode: Set envNode = objXML.SelectNodes(xPath)
	
	'Return xml object (With Set command)
	Set GetTestEnvConfigurationObject = envNode.item(0)

End Function

' Fetch any key value from the test environemnt.
Function GetXMLChildNodeValue(parentXMLNode, Key)

	Dim xpath: xpath = "./"&Key
	Dim childXMLNode: Set childXMLNode = parentXMLNode.SelectNodes(xpath)
	
	'Return string value (Without Set command)
	GetXMLChildNodeValue = childXMLNode.item(0).nodeTypedValue

End Function