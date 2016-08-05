Sub CreateMenuItem()
'VBA546
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Adds two menuitems to menu "Delete objects"
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete Rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete Circles")
End Sub
