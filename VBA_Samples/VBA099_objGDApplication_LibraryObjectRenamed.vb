Private Sub objGDApplication_LibraryObjectRenamed(ByVal LibObject As IHMIFolderItem, ByVal OldName As String)
'VBA99
    MsgBox "The object " & OldName & " is renamed in: " & LibObject.DisplayName
End Sub