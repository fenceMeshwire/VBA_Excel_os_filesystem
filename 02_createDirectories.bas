Option Explicit

Sub createDirectory()

Dim strParentDirectory As String
Dim strDirectory As String
Dim strResultDir As String
Dim strCheckDir As String

strParentDirectory = "C:\Users\...\"
strDirectory = "TestDirectory"
strResultDir = strParentDirectory & strDirectory

strCheckDir = Dir(strResultDir, vbDirectory)

If strCheckDir = "" Then
  MkDir strResultDir
Else
  Debug.Print "Directory already exists"
End If

End Sub
