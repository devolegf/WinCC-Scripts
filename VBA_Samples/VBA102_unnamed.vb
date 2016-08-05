Private Sub objGDApplication_NewLibraryFolder(ByVal LibObject As IHMIFolderItem)
'VBA102
    MsgBox "The library-folder " & LibObject.DisplayName & " was added."
End Sub
