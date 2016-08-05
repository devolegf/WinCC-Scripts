Private Sub objGDApplication_LibraryFolderRenamed(ByVal LibObject As HMIFolderItem, ByVal OldName As String)
'VBA98
    MsgBox "The Library-folder " & OldName & " is renamed in: " & LibObject.DisplayName
End Sub