Private Sub Document_LibraryObjectAdded(ByVal LibObject As IHMIFolderItem, CancelForwarding As Boolean)
'VBA100
    Dim strObjName As String
'
    '"strObjName" contains the name of the added object
    strObjName = LibObject.DisplayName
    MsgBox "Object " & strObjName & " was added to the picture."
End Sub