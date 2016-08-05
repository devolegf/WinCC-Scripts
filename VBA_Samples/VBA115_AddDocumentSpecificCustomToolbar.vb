Sub AddDocumentSpecificCustomToolbar()
'VBA115
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
 
    Set objToolbar = ActiveDocument.CustomToolbars.Add("DocToolbar")
    
    'Add toolbar-items to the userdefined toolbar
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "Mein erstes Symbol-Icon")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(3, "tItem1_3", "Mein zweites Symbol-Icon")
'
    'Insert seperatorline between the two tollbaritems
    Set objToolbarItem = objToolbar.ToolbarItems.InsertSeparator(2, "tSeparator1_2")
End Sub
