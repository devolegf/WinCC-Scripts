Sub CreateToolbar()
'VBA533
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Dim strFileWithPath
    Set objToolbar = ActiveDocument.CustomToolbars.Add("Tool1_1")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "ti1_1", "myFirstToolbaritem")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(2, "ti1_2", "mySecondToolbaritem")
'
    'ITo use this example copy a *.ICO-Graphic
    'to the "GraCS"-Folder of the actual project.
    'Replace the filename "EZSTART.ICO" in the next commandline
    'with the name of the ICO-Graphic you copied
    strFileWithPath = Application.ApplicationDataPath & "EZSTART.ICO"
'
    'To assign the symbol-icon to the first toolbaritem
    objToolbar.ToolbarItems(1).Icon = strFileWithPath
End Sub
