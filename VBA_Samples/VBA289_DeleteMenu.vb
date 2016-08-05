Sub DeleteMenu()
'VBA289
    Dim objMenu As HMIMenu
    Set objMenu = ActiveDocument.CustomMenus(1)
    objMenu.Delete
End Sub
