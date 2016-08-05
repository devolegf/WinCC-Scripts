Sub CreateMenuItem()
'VBA366
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Create new menu "Delete Objects":
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete Objects")
'
    'Add two menuitems to the menu "Delete Objects
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete Rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete Circles")
End Sub
