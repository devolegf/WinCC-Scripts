Sub CreateDocumentMenus()
'VBA728
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    'Add menu:
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
'
    'Add menuitems to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "My first menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second menuitem")
'
    'Add seperator to custom-menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    'Add submenu to custom-menu:
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "My first submenu")
'
    'Add menuitems to submenu:
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "My first submenuitem")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "My second submenuitem")
'
    'Assign statustexts to every menuitem
    With objDocMenu
        .MenuItems(1).StatusText = "My first menuitem"
        .MenuItems(2).StatusText = "My second menuitem"
        .MenuItems(4).SubMenu.Item(1).StatusText = "My first submenuitem"
        .MenuItems(4).SubMenu.Item(2).StatusText = "My second submenuitem"
    End With
End Sub
