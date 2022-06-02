Option Explicit

Sub getDirectoryNamesInParentDirectory()

Dim intCounter As Integer
Dim strParentDirectory As String
Dim strDirName As String
Dim wksSheet As Worksheet

Set wksSheet = Sheet1   ' Set Worksheet for output of directory information

strParentDirectory = "C:\Users\...\"   ' Set the directory path of the parent directory
strDirName = Dir(strParentDirectory, vbDirectory)

intCounter = 0

Do While strDirName <> ""
  strDirName = Dir
  If Not InStr(strDirName, ".") > 0 Then
    intCounter = intCounter + 1
    wksBlatt.Cells(intCounter, 1).Value = strDirName
  End If
Loop

End Sub
