Option Explicit

Sub remove_directory()

Dim strPath As String

strPath = "C:\Users\...\remove_directory"

On Error GoTo file_error

If Dir(strPath, vbDirectory) <> "" Then
  RmDir strPath
ElseIf Dir(strPath, vbDirectory) = "" Then
  MsgBox ("Remove directory not possible. Directory does not exist.")
End If

file_error:
  MsgBox ("Remove directory not possible. Directory contains file(s).")
  
End Sub
