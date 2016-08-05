Sub LDTooltipInfo()
'VBA597
    Dim colLangTexts As HMILanguageTexts
    Dim objLangText As HMILanguageText
    Dim iAnswer As Integer
    Set colLangTexts = ActiveDocument.CustomToolbars(1).ToolbarItems(1).LDTooltipTexts
    For Each objLangText In colLangTexts
        iAnswer = MsgBox(objLangText.DisplayName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLangText
End Sub
