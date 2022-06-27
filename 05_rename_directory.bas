Sub rename_directory()

Dim fso As Object
Dim strOldName, strNewName As String

Set fso = CreateObject("Scripting.FileSystemObject")

strOldName = "C:\temp\old"
strNewName = "C:\temp\new"

fso.Movefolder strOldName, strNewName

End Sub
