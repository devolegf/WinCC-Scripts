Sub AddStatusAndTooltipTextsToAppToolbar1()
'VBA9
    Dim objToolbar1 As HMIToolbar
'
    'Variable "StatusTextToolbarItem1" for foreign statustexts
    Dim objStatusTextToolbarItem1 As HMILanguageText
'
    'Variable "TooltipTextToolbarItem1 for foreign tooltiptexts
    Dim objTooltipTextToolbarItem1 As HMILanguageText
 
    Set objToolbar1 = Application.CustomToolbars("AppToolbar1")
'
    'Assign a statustext to a toolbaritem:
    objToolbar1.ToolbarItems("tItem1_1").StatusText = "Statustext für das erste Symbol-Icon"
'
    'Assign a foreign statustext to a toolbaritem:
    Set objStatusTextToolbarItem1 = objToolbar1.ToolbarItems("tItem1_1").LDStatusTexts.Add(1033, "This is my first status text in english")
'
    'Assign a foreign tooltiptext to a toolbaritem:
    Set objTooltipTextToolbarItem1 = objToolbar1.ToolbarItems("tItem1_1").LDTooltipTexts.Add(1033, "This is my first tooltip text in english")
End Sub