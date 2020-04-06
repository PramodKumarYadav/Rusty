'To avoid errors due to typos in variable names
Option Explicit
 
Function LoadAllFunctionLibraries(pathFnLibraryDir)
	
	' Get all the files in the pathFnLibraryDir 
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim objRootFolder: Set objRootFolder = objFSO.GetFolder(pathFnLibraryDir)

	Dim objFiles: Set objFiles = objRootFolder.Files
	Dim objFile
	For Each objFile in objFiles
		'msgbox objFile.Name 
		'If file is of type vbs, then load this file.
		If LCase(objFSO.GetExtensionName(objFile.Name)) = "vbs" Then
            'Associate the function library
			LoadFunctionLibrary objFile.Path 
        End If		
	Next

	Call ShowSubfolders(objRootFolder)
	
End Function

Sub ShowSubFolders(objRootFolder)
	
	Dim objSubfolder
    For Each objSubfolder in objRootFolder.SubFolders
        
		Dim objChildFolder: Set objChildFolder = objFSO.GetFolder(objSubfolder.Path)
        
		Dim objFiles: Set objFiles = objChildFolder.Files 
		
		Dim objFile
		For Each objFile in objFiles
			'msgbox objFile.Name 
            'If file is of type vbs, then load this file.
			If LCase(objFSO.GetExtensionName(objFile.Name)) = "vbs" Then
				'Associate the function library
				LoadFunctionLibrary objFile.Path 
			End If
        Next

        Call ShowSubFolders(objSubfolder)	
    Next

End Sub