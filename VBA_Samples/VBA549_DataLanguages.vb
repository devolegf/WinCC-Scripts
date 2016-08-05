Sub DataLanguages()
'VBA549
    Dim colDataLang As HMIDataLanguages
    Dim objDataLang As HMIDataLanguage
    Dim nLangID As Long
    Dim strLangName As String
    Dim iAnswer As Integer
    Set colDataLang = Application.AvailableDataLanguages
    For Each objDataLang In colDataLang
        nLangID = objDataLang.LanguageID
        strLangName = objDataLang.LanguageName
        iAnswer = MsgBox(nLangID & " " & strLangName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objDataLang
End Sub
