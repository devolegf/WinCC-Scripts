Sub InsertToolbarItems()
'VBA7
    Dim objToolbar1 As HMIToolbar
    Dim objToolbarItem1 As HMIToolbarItem
'
    'Add a new toolbar:
    Set objToolbar1 = Application.CustomToolbars.Add("AppToolbar1")
    'Adds two toolbar-items to the toolbar
    '("InsertToolbarItem(Position, Key, DefaultToolTipText)"-Methode):
    Set objToolbarItem1 = objToolbar1.ToolbarItems.InsertToolbarItem(1, "tItem1_1", "First Symbol-Icon")
    Set objToolbarItem1 = objToolbar1.ToolbarItems.InsertToolbarItem(3, "tItem1_2", "Second Symbol-Icon")
'
    'Adds a seperator between the two toolbar-items
    '("InsertSeparator(Position, Key)"-Methode):
    Set objToolbarItem1 = objToolbar1.ToolbarItems.InsertSeparator(2, "tSeparator1_3")
End Sub