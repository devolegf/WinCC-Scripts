Sub CreateApplicationMenus()
'VBA3
    'Declaration of menus...:
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    '
    'Add menus. Parameters are "Position", "Key" und "DefaultLabel":
    Set objMenu1 = Application.CustomMenus.InsertMenu(1, "AppMenu1", "App_Menu_1")
    Set objMenu2 = Application.CustomMenus.InsertMenu(2, "AppMenu2", "App_Menu_2")
End Sub