Sub CreateMenuItem()
'VBA640
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim iIndex As Integer
    iIndex = 1
'
    'Add new menu "Delete objects" to the menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to menu "Delete objects"
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Output label of menu:
    MsgBox ActiveDocument.CustomMenus(1).Label
'
    'Output labels of all menuitems:
    For iIndex = 1 To objMenu.MenuItems.Count
        MsgBox objMenu.MenuItems(iIndex).Label
    Next iIndex
End Sub
