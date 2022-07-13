Option Explicit

Sub remove_directory()

Dim strPath As String

strPath = "C:\Users\username\directory"

On Error GoTo file_error

If Dir(strPath, vbDirectory) <> "" Then
  RmDir strPath
ElseIf Dir(strPath, vbDirectory) = "" Then
  MsgBox ("Remove directory not possible. Directory does not exist.")
End If

file_error:
  MsgBox ("Remove directory not possible. " & _
    "Directory contains further directories or files.")
  
End Sub
