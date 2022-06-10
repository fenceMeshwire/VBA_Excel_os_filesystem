Sub map_directory()

Dim objFSO As Object
Dim objDirectory As Object
Dim objFile As Object
Dim objFileText As Object
Dim strFile As String
Dim strTextFile As String
Dim strPathTextFile As String
Dim strPath As String

strPath = "C:\Users\name"
strTextFile = "Outfile.txt"
strPathTextFile = strPath & "\" & strTextFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objDirectory = objFSO.getfolder(strPath)

If Dir(strPathTextFile, vbDirectory) <> "" Then
  Exit Sub
End If

Open strTextFile For Output As #1

For Each objFile In objDirectory.Files
  strFile = CStr(objFile.Name)
  Print #1, strFile
Next objFile

Close #1

End Sub
