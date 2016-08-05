Sub ShowFirstObjectOfCollection()
'VBA348
    Dim strName As String
    strName = ActiveDocument.CustomToolbars(1).Key
    MsgBox strName
End Sub
