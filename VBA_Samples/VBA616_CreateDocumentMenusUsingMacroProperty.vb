Sub CreateDocumentMenusUsingMacroProperty()
'VBA616
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Set objDocMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "dmItem1_1", "My first menuitem")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "dmItem1_2", "My second menuitem")
'
    'To assign a macro to every menuitem:
    With ActiveDocument.CustomMenus("DocMenu1")
        .MenuItems("dmItem1_1").Macro = "TestMacro1"
        .MenuItems("dmItem1_2").Macro = "TestMacro2"
    End With
End Sub


Sub TestMacro1()
    MsgBox "TestMacro1 is executed"
End Sub
 

Sub TestMacro2()
    MsgBox "TestMacro2 is executed"
End Sub
