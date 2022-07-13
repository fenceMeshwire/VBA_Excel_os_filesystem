Option Explicit

Sub list_files_in_dir()

Dim strFile As String
Dim strPath As String

strPath = "C:\Users\user\directory\"

strFile = Dir(strPath)

While strFile <> ""
  Debug.Print strFile, FileDateTime(strPath & strFile)
  strFile = Dir
Wend

End Sub
