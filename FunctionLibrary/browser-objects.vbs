'Reference & Credits:
'First things first, I based my work on web automation (this script) from George.Zhao): https://github.com/persistz/WebTesting_Common_Function_for_UFT_QTP
'All Credits and kudos to him for being the only perons I found on GitHub to share his work on UFT.

'To avoid typo errors due to variable spelling mistakes
Option Explicit

Function GetBrowserObjectWithHomePageUnknown()
	
	'Set object based on the property browser name
	Dim objBrowser: Set objBrowser = Browser("micclass:=Browser")

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objBrowser, "HomePage Unknown", "Browser" 

	'Assign this object to function
	Set GetBrowserObjectWithHomePageUnknown = objBrowser

	'Now release this object memory
	Set objBrowser = Nothing
	
End Function

Function GetBrowserObject(name)

	'Set object based on the property browser name
	Dim objBrowser: Set objBrowser = Browser("name:="&name)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objBrowser, name, "Browser" 

	'Assign this object to function
	Set GetBrowserObject = objBrowser

	'Now release this object memory
	Set objBrowser = Nothing
	
End Function

Function GetPageObject(name, title)
	
	'Set object based on the property browser name and page title
	Dim objPage: Set objPage = Browser("name:="&name).Page("title:="&title)	

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objPage, title, "Page" 

	'Assign this object to function
	Set GetPageObject = objPage

	'Now release this object memory
	Set objPage = Nothing
	
End Function

Function CheckIfObjectExistsAndIsVisible (objToCheck, objPropertyValue, objType)
	
	'Check if object exists
	CheckIfObjectExists objToCheck, objPropertyValue, objType
	
	'Check if object is visible
	CheckIfObjectIsVisible objToCheck, objPropertyValue, objType
	
End Function

Function CheckIfObjectExists(objToCheck, objPropertyValue, objType )
	
	' Set status to fail and exit test if object doesnt exist 
	' Do not add status for pass (using micDone), since UFT by default will add it for below check i.e. objToCheck.exist (adding it will result in duplicate statements)	

	' Report the result	
	Dim testStep: testStep = "Find object type ["&objType&"] with property value ["&objPropertyValue&"]"
	If objToCheck.exist Then		
		Call SetResultRecordForTestReport(testStep, "Found","Found", "Pass")			
	Else
		Call SetResultRecordForTestReport(testStep, "Not Found","Found", "Fail")
		Call CreateReportAndExitTests()
	End if

End Function

Function CheckIfObjectIsVisible (objToCheck, objPropertyValue, objType)
	
	' Check if object is visible (sometimes the object exists but is not visible and thus the operations will fail).	
	' Do not add status for pass (using micDone), since UFT by default will add it for below check i.e. objToCheck.WaitProperty (adding it will result in duplicate statements)	

	' Report the result	
	Dim testStep: testStep = "Check if Object type ["&objType&"] and property value ["&objPropertyValue&"] is visible"
	If objToCheck.WaitProperty ("visible", "True", 10000) Then
		Call SetResultRecordForTestReport(testStep, "Visible","Visible", "Pass")		
	Else
		Call SetResultRecordForTestReport(testStep, "Not Visible","Visible", "Fail")
		Call CreateReportAndExitTests()
	End if

End Function

Function GetFrameObject(objParent, name)
	
	'Set object based on the parent object and property name
	Dim objPage: Set objPage = objParent.Frame("name:="&name)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objPage, name, "Frame" 

	'Assign this object to function
	Set GetFrameObject = objPage

	'Now release this object memory
	Set objPage = Nothing
	
End Function

Function GetWebTable(objParent, html_id, xpath)
	
	'Set object based on the parent object and property html id and xpath
	Dim objWebTable: Set objWebTable = objParent.WebTable("html id:="&html_id, "xpath:="&xpath)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objWebTable, html_id, "WebTable" 

	'Assign this object to function
	Set GetWebTable = objWebTable

	'Now release this object memory
	Set objWebTable = Nothing
	
End Function

Function GetLink(objParent, outertext)
	
	'Set object based on the parent object and property outertext
	Dim objLink: Set objLink = objParent.Link("outertext:="&outertext)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objLink, outertext, "Link" 

	'Assign this object to function
	Set GetLink = objLink

	'Now release this object memory
	Set objLink = Nothing
	
End Function

Function GetWebElement(objParent, outertext, xpath)
	
	'Set object based on the parent object and property outertext and xpath
	Dim objWebElement: Set objWebElement = objParent.WebElement("outertext:="&outertext, "xpath:="&xpath)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objWebElement, outertext, "WebElement" 

	'Assign this object to function
	Set GetWebElement = objWebElement

	'Now release this object memory
	Set objWebElement = Nothing
	
End Function

Function GetImage(objParent, title, xpath)
	
	'Set object based on the parent object and property title and xpath
	Dim objImage: Set objImage = objParent.Image("title:="&title, "xpath:="&xpath)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objImage, title, "Image" 

	'Assign this object to function
	Set GetImage = objImage

	'Now release this object memory
	Set objImage = Nothing
	
End Function

Function GetWebEdit (objParent, name)
	
	'Set object based on the parent object and property name
	Dim objWebEdit: Set objWebEdit = objParent.WebEdit("name:="&name)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objWebEdit, name, "WebEdit" 

	'Assign this object to function
	Set GetWebEdit = objWebEdit

	'Now release this object memory
	Set objWebEdit = Nothing
	
End Function

Function GetWebButton (objParent, name )
	
	'Set object based on the parent object and property name
	Dim objWebButton: Set objWebButton = objParent.WebButton("name:="&name)

	'Check and Continue only if the object exists and is visible
	CheckIfObjectExistsAndIsVisible objWebButton, name, "WebButton" 

	'Assign this object to function
	Set GetWebButton = objWebButton

	'Now release this object memory
	Set objWebButton = Nothing
	
End Function
