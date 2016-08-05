Sub ShowFirstMenuOfMenucollection()
'VBA288
    Dim strName As String
    strName = ActiveDocument.CustomMenus(1).Label
    MsgBox strName
End Sub
