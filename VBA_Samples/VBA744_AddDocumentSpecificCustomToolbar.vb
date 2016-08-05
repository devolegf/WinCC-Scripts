Sub AddDocumentSpecificCustomToolbar()
'VBA744
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Set objToolbar = ActiveDocument.CustomToolbars.Add("DocToolbar")
'
    'Add symbol-icon to userdefined toolbar
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "My first symbol-icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(3, "tItem1_3", "My second symbol-icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertSeparator(2, "tSeparator1_2")
End Sub
