Sub CreateMenuItem()
'VBA548
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim iIndex As Integer
    iIndex = 1
'
    'Add new menu "Delete objects" to menubar
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Adds two menuitems to menu "Delete objects"
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
    MsgBox ActiveDocument.CustomMenus(1).Label
    For iIndex = 1 To objMenu.MenuItems.Count
        MsgBox objMenu.MenuItems(iIndex).Label
    Next iIndex
End Sub
