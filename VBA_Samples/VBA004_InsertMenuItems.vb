Sub InsertMenuItems()
'VBA4
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    Dim objMenuItem1 As HMIMenuItem
    Dim objSubMenu1 As HMIMenuItem
    'Create Menu:
    Set objMenu1 = Application.CustomMenus.InsertMenu(1, "AppMenu1", "App_Menu_1")
    
    'Next lines add menu-items to userdefined menu.
    'Parameters are "Position", "Key" and DefaultLabel:
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(1, "mItem1_1", "App_MenuItem_1")
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(2, "mItem1_2", "App_MenuItem_2")
'
    'Adds seperator to menu ("Position", "Key")
    Set objMenuItem1 = objMenu1.MenuItems.InsertSeparator(3, "mItem1_3")
'
    'Adds a submenu into a userdefined menu
    Set objSubMenu1 = objMenu1.MenuItems.InsertSubMenu(4, "mItem1_4", "App_SubMenu_1")
 '
    'Adds a menu-item into a submenu
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(5, "mItem1_5", "App_SubMenuItem_1")
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(6, "mItem1_6", "App_SubMenuItem_2")
End Sub
