Sub InsertDocumentSpecificMenu()
'VBA292
    Dim objMenu As HMIMenu
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "d_Menu1", "myDocumentMenu")
End Sub
