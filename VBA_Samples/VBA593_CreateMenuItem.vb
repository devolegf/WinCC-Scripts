Sub CreateMenuItem()
'VBA593
    Dim objMenu As HMIMenu
    Dim objMenuItem1 As HMIMenuItem
    Dim objMenuItem2 As HMIMenuItem
    Dim objLangStateText As HMILanguageText
'
    'Add new menu "Delete objects" to menubar:
    Set objMenu = ActiveDocument.CustomMenus.InsertMenu(1, "DeleteObjects", "Delete objects")
'
    'Add two menuitems to the new menu
    Set objMenuItem1 = objMenu.MenuItems.InsertMenuItem(1, "DeleteAllRectangles", "Delete rectangles")
    Set objMenuItem2 = objMenu.MenuItems.InsertMenuItem(2, "DeleteAllCircles", "Delete circles")
'
    'Define foreign-language labels for menuitem "Delete rectangles":
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1033, "English_Delete rectangles")
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1032, "Greek_Delete rectangles")
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1034, "Spanish_Delete rectangles")
    Set objLangStateText = objMenuItem1.LDStatusTexts.Add(1036, "French_Delete rectangles")
End Sub
