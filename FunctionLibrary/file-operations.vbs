'Reference & Credits: https://admhelp.microfocus.com/uft/en/all/CodeSamplesPlus_Help/Content/Code_Samples_Plus/CSP_DebuggingWFileOps.htm

' Creates a specified file and returns a TextStream object that can be used to read from or write to the file.
' Example of usage:
' Set f = CreateFile("d:	emp\beenhere.txt", True)
' f.WriteLine Now
' f.Close
Function CreateFile(sFilename, bOverwrite)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set CreateFile = fso.CreateTextFile(sFilename, bOverwrite)
End Function

' Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.
' iomode: 1 - ForReading, 2 - ForWriting, 8 - ForAppending
' Example of usage
' Set f = OpenFile("d:	emp\beenhere.txt", 2, True)
' f.WriteLine Now
' f.Close
Function OpenFile(sFilename, iomode, create)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set OpenFile = fso.OpenTextFile(sFilename, iomode, create)
End Function

' Appends a line to a file.
' Example of usage:
' AppendToFile "d:	emp\beenhere.txt", Now
Function AppendToFile(sFilename, sLine)
	Const ForAppending = 8
	If sFilename = "" Then
		sFilename = Environment("SystemTempDir") & "\QTDebug.txt"
	End If
	Set f = OpenFile(sFilename, ForAppending, True)
	f.WriteLine sLine
	f.Close
End Function

' Writes a line to a file.
' Destroys the current content of the file.
' Example of usage:
' WriteToFile "d:	emp\beenhere.txt", Now
Function WriteToFile(sFilename, sLine)
	Const ForWriting = 2
	If sFilename = "" Then
		sFilename = Environment("SystemTempDir") & "\QTDebug.txt"
	End If
	Set f = OpenFile(sFilename, ForWriting, True)
	f.WriteLine sLine
	f.Close
End Function