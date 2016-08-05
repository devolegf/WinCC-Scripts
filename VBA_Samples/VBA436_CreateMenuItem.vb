Sub CreateMenuItem()
'VBA436
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete Rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete Circles")
    With objMenu.MenuItems
        .Item("DeleteAllRectangles").Checked = True
    End With
End Sub
