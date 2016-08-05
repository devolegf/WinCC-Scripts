Sub InsertMenuItems()
'VBA5
    'Execute this procedure first
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    Dim objMenuItem1 As HMIMenuItem
    Dim objSubMenu1 As HMIMenuItem
    'Add Menu:
    Set objMenu1 = Application.CustomMenus.InsertMenu(1, "AppMenu1", "App_Menu_1")
    
    'Next lines add menu-items to userdefined menu.
    'Parameters are "Position", "Key" and DefaultLabel:
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(1, "mItem1_1", "App_MenuItem_1")
    Set objMenuItem1 = objMenu1.MenuItems.InsertMenuItem(2, "mItem1_2", "App_MenuItem_2")
'
    'Adds seperator to menu ("Position", "Key")
    Set objMenuItem1 = objMenu1.MenuItems.InsertSeparator(3, "mItem1_3")
'
    'Adds a submenu to a userdefined menu
    Set objSubMenu1 = objMenu1.MenuItems.InsertSubMenu(4, "mItem1_4", "App_SubMenu_1")
 '
    'Adds a menu-item to a submenu
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(5, "mItem1_5", "App_SubMenuItem_1")
    Set objMenuItem1 = objSubMenu1.SubMenu.InsertMenuItem(6, "mItem1_6", "App_SubMenuItem_2")
End Sub

Sub MultipleLanguagesForAppMenu1()
    'Execute this procedure after execution of "InsertMenuItems()" 

    'Object "objLanguageTextMenu1" contains the
    'foreign-language labels for the menu
    Dim objLanguageTextMenu1 As HMILanguageText
'
    'Object "objLanguageTextMenu1Item" contains the
    'foreign-language labels for the menu-items
    Dim objLanguageTextMenuItem1 As HMILanguageText
 
    Dim objMenu As HMIMenu
    Dim objSubMenu1 As HMIMenuItem
 
    Set objMenu1 = Application.CustomMenus("AppMenu1")
    Set objSubMenu1 = Application.CustomMenus("AppMenu1").MenuItems("mItem1_4")
'
    'Ads foreign-language label into a menu:
    '("Add(LCID, DisplayName)"-Methode:
    Set objLanguageTextMenu1 = objMenu1.LDLabelTexts.Add(1033, "English_App_Menu_1")
'
    'Adds foreign-language label into a menuitem:
    Set objLanguageTextMenuItem1 = objMenu1.MenuItems("mItem1_1").LDLabelTexts.Add(1033, "My first menu item")
'
    'Adds a foreign-language label into a submenu:
    Set objLanguageTextMenuItem1 = objSubMenu1.SubMenu.Item("mItem1_5").LDLabelTexts.Add(1033, "My first submenu item")
End Sub