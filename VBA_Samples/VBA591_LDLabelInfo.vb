Sub LDLabelInfo()
'VBA591
    Dim colLangTexts As HMILanguageTexts
    Dim objLangText As HMILanguageText
    Dim iAnswer As Integer
'
    'Save all labels of menu into collection "colLangTexts":
    Set colLangTexts = ActiveDocument.CustomMenus("DeleteObjects").LDLabelTexts
    For Each objLangText In colLangTexts
        iAnswer = MsgBox(objLangText.DisplayName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLangText
End Sub
