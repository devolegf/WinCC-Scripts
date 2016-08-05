Sub CreateMenuItem()
'VBA590
    Dim objMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    Dim objLangText As HMILanguageText
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Define foreign-language labels for menu "Delete objects":
    Set objLangText = objMenu.LDLabelTexts.Add(1033, "English_Delete objects")
    Set objLangText = objMenu.LDLabelTexts.Add(1032, "Greek_Delete objects")
    Set objLangText = objMenu.LDLabelTexts.Add(1034, "Spanish_Delete objects")
    Set objLangText = objMenu.LDLabelTexts.Add(1036, "French_Delete objects")
End Sub
