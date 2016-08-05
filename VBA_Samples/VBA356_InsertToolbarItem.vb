Sub InsertToolbarItem()
'VBA356
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Set objToolbar = ActiveDocument.CustomToolbars.Add("d_Toolbar2")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "t_Item2_1", "ToolbarItem 1")
End Sub
