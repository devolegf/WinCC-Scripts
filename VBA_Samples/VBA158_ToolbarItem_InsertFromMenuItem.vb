Sub ToolbarItem_InsertFromMenuItem()
'VBA158
    Dim objMenu As HMIMenu
    Dim objToolbarItem As HMIToolbarItem
    Dim objToolbar As HMIToolbar
    Dim objMenuItem As HMIMenuItem
    Set objMenu = Application.CustomMenus.InsertMenu(1, "Menu1", "TestMenu")
'
    '*************************************************
    '* Note:
    '* The object-reference has to be unique.
    '*************************************************
'
    Set objMenuItem = Application.CustomMenus(1).MenuItems.InsertMenuItem(1, "MenuItem1", "Hello World")
    Application.CustomMenus(1).MenuItems(1).Macro = "HelloWorld"
    Set objToolbar = Application.CustomToolbars.Add("Toolbar1")
    Set objToolbarItem = Application.CustomToolbars(1).ToolbarItems.InsertFromMenuItem(1, "ToolbarItem1", objMenuItem, "Call's Hello World of TestMenu")
End Sub

Sub HelloWorld()
    MsgBox "Procedure 'HelloWorld()' is execute."
End Sub
