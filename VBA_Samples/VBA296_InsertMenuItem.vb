Sub InsertMenuItem()
'VBA296
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(2, "d_Menu2", "DocMenu2")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "m_Item2_1", "MenuItem 1")
End Sub
