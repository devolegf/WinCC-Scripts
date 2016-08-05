Sub DisableMenuItem()
'VBA486
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
'
    'Add a new menu "Delete objects"
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Disable menuitem "Delete circles"
    With ActiveDocument.CustomMenus("DeleteObjects").MenuItems("DeleteAllCircles")
        .Enabled = False
    End With
End Sub
