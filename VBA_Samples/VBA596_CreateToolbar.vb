Sub CreateToolbar()
'VBA596
    Dim objToolbar As HMIToolbar
    Dim objToolbarItem As HMIToolbarItem
    Dim objLangText As HMILanguageText
    Dim strFileWithPath
'
    'Create toolbar with two toolbar-items:
    Set objToolbar = ActiveDocument.CustomToolbars.Add("Tool1_1")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(1, "ti1_1", "myFirstToolbaritem")
    Set objToolbarItem = objToolbar.ToolbarItems.InsertToolbarItem(2, "ti1_2", "mySecondToolbaritem")
'
    'In order that the example runs correct copy a *.ICO-Graphic
    'into the "GraCS"-Folder of the actual project.
    'Replace the filename "EZSTART.ICO" in the next commandline
    'with the name of the ICO-Graphic you copied
    strFileWithPath = Application.ApplicationDataPath & "EZSTART.ICO"
'
'
    'To assign the symbol-icon to the first toolbaritem
    objToolbar.ToolbarItems(1).Icon = strFileWithPath
'
    'Define foreign-language tooltiptexts
    Set objLangText = objToolbar.ToolbarItems(1).LDTooltipTexts.Add(1036, "French_Tooltiptext")
    Set objLangText = objToolbar.ToolbarItems(1).LDTooltipTexts.Add(1034, "Spanish_Tooltiptext")
End Sub
