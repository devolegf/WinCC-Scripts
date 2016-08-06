Private Sub objGDApplication_BeforeLibraryFolderDelete(ByVal LibObject As HMIFolderItem, Cancel As Boolean)
'VBA81
    MsgBox "The library-folder " & LibObject.Name & " will be delete..."
End Sub