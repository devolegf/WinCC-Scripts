Sub CreateDocumentMenus()
'VBA25
    'Declare menuobjects:
    Dim objMenu1 As HMIMenu
    Dim objMenu2 As HMIMenu
    
    'Insert Menus ("InsertMenu"-Methode) with
    'Parameters - "Position", "Key", "DefaultLabel":
    Set objMenu1 = ActiveDocument.CustomMenus.InsertMenu(1, "DocMenu1", "Doc_Menu_1")
    Set objMenu2 = ActiveDocument.CustomMenus.InsertMenu(2, "DocMenu2", "Doc_Menu_2")
End Sub