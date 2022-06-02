Option Explicit

Sub createFile()

Dim strDirectory As String
Dim strFile As String
Dim strResultPath As String
Dim strCheckDir As String
Dim objFileSystem As Object
Dim objFile As Object

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

strDirectory = "C:\Users\...\"
strFile = "TestFile" ' Note: No file extension.
strResultPath = strDirectory & strFile

strCheckDir = Dir(strResultPath, vbDirectory)

If strCheckDir = "" Then
  Set objFile = objFileSystem.CreateTextFile(strResultPath)
Else
  Debug.Print "File already exists."
End If

Set objFileSystem = Nothing
Set objFile = Nothing

End Sub
