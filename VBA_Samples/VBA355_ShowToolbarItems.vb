Sub ShowToolbarItems()
'VBA355
    Dim colToolbarItems As HMIToolbarItems
    Dim objToolbarItem As HMIToolbarItem
    Dim strTypeList As String
    Set colToolbarItems = ActiveDocument.CustomToolbars(1).ToolbarItems
    If 0 <> colToolbarItems.Count Then
        For Each objToolbarItem In colToolbarItems
            strTypeList = strTypeList & objToolbarItem.ToolbarItemType & vbCrLf
        Next objToolbarItem
    Else
        strTypeList = "No Toolbaritems existing"
    End If
    MsgBox strTypeList
End Sub
