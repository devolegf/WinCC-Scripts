Sub CreateDocumentMenus()
'VBA719
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objSubMenu As HMIMenuItem
'
    'Add menu to menubar:
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    '
    'Add menuitems to the new menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "&My first Menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second Menuitem")
'
    'Add seperator to menu:
    Set objMenuItem = objDocMenu.MenuItems.InsertSeparator(3, "dSeparator1_3")
'
    'Add submenu to the menu:
    Set objSubMenu = objDocMenu.MenuItems.InsertSubMenu(4, "dSubMenu1_4", "My first submenu")
'
    'Add menuitems to the submenu:
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(5, "dmItem1_5", "My first submenuitem")
    Set objMenuItem = objSubMenu.SubMenu.InsertMenuItem(6, "dmItem1_6", "My second submenuitem")
'
    ActiveDocument.CustomMenus("DocMenu1").MenuItems(1).ShortCut = "STRG+SHIFT+M"
End Sub
