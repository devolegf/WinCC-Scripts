Sub InsertApplicationSpecificMenu()
'VBA291
    Dim objMenu As HMIMenu
    Set objMenu = Application.CustomMenus.InsertMenu(1, "a_Menu1", "myApplicationMenu")
End Sub
