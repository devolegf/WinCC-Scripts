Sub AddDocumentSpecificCustomToolbar()
'VBA161
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
 
    Set objToolbar = ActiveDocument.CustomToolbars.Add("DocToolbar")
    
    'Add toolbar-item to userdefined toolbar
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "First symbol-icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(3, "tItem1_3", "Second symbol-icon")
'
    'Insert dividing rule between first and second symbol-icon
    Set objToolbarItem = objToolbar.ToolbarItems.InsertSeparator(2, "tSeparator1_2")
End Sub
