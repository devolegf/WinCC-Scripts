Sub ShowMenuItems()
'VBA295
    Dim colMenuItems As HMIMenuItems
    Dim objMenuItem As HMIMenuItem
    Dim strItemList As String
    Set colMenuItems = ActiveDocument.CustomMenus(1).MenuItems
    For Each objMenuItem In colMenuItems
        strItemList = strItemList & objMenuItem.Label & vbCrLf
    Next objMenuItem
    MsgBox strItemList
End Sub
