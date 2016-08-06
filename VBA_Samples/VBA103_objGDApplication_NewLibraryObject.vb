Private Sub objGDApplication_NewLibraryObject(ByVal LibObject As IHMIFolderItem)
'VBA103
    MsgBox "The object " & LibObject.DisplayName & " was added."
End Sub