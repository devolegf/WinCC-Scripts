Private Sub objGDApplication_BeforeLibraryObjectDelete(ByVal LibObject As HMIFolderItem, Cancel As Boolean)
'VBA82
    MsgBox "The object " & LibObject.Name & " will be delete..."
End Sub