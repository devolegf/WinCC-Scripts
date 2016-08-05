Sub ShowMenuTypes()
'VBA641
    Dim iMaxMenuItems As Integer
    Dim iMenuItemType As Integer
    Dim strMenuItemType As String
    Dim iIndex As Integer
    iMaxMenuItems = ActiveDocument.CustomMenus(1).MenuItems.Count
    For iIndex = 1 To iMaxMenuItems
        iMenuItemType = ActiveDocument.CustomMenus(1).MenuItems(iIndex).MenuItemType
        Select Case iMenuItemType
            Case 0
                strMenuItemType = "Trennstrich (Separator)"
            Case 1
                strMenuItemType = "Untermenü (SubMenu)"
            Case 2
                strMenuItemType = "Menüeintrag (MenuItem)"
        End Select
        MsgBox iIndex & ". Menuitemtype: " & strMenuItemType
    Next iIndex
End Sub
