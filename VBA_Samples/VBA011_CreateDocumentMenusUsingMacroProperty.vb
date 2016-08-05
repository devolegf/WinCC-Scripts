Sub CreateDocumentMenusUsingMacroProperty()
'VBA11
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "First Menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "Second Menuitem")
'
    'Assign a VBA-macro to every menu item
    With ActiveDocument.CustomMenus("DocMenu1")
        .MenuItems("dmItem1_1").Macro = "TestMacro1"
        .MenuItems("dmItem1_2").Macro = "TestMacro2"
    End With
End Sub