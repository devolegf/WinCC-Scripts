Sub ShowLabelTexts()
'VBA481
    Dim objLangText As HMILanguageText
    Dim iIndex As Integer
    For iIndex = 1 To ActiveDocument.CustomMenus(1).LDLabelTexts.Count
        Set objLangText = ActiveDocument.CustomMenus(1).LDLabelTexts(iIndex)
        MsgBox objLangText.DisplayName
    Next iIndex
End Sub
