Sub ShowFirstObjectOfCollection()
'VBA293
    Dim strName As String
    strName = ActiveDocument.CustomMenus(1).MenuItems(1).Label
    MsgBox strName
End Sub
