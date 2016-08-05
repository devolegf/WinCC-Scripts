Sub ShowPositionOfCustomMenuItems()
'VBA683
    Dim objMenu As HMIMenu
    Dim iMaxMenuItems As Integer
    Dim iPosition As Integer
    Dim iIndex As Integer
    Set objMenu = ActiveDocument.CustomMenus(1)
    iMaxMenuItems = objMenu.MenuItems.Count
    For iIndex = 1 To iMaxMenuItems
        iPosition = objMenu.MenuItems(iIndex).Position
        MsgBox "Position of the " & iIndex & ". menuitem: " & iPosition
    Next iIndex
End Sub
