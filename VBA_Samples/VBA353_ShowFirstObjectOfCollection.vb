Sub ShowFirstObjectOfCollection()
'VBA353
    Dim strType As String
    strType = ActiveDocument.CustomToolbars(1).ToolbarItems(1).ToolbarItemType
    MsgBox strType
End Sub
