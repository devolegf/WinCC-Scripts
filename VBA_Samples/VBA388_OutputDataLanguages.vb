Sub OutputDataLanguages()
'VBA388
    Dim colDataLang As HMIDataLanguages
    Dim objDataLang As HMIDataLanguage
    Dim strLangList As String
    Dim iCounter As Integer
'
    'Save collection of datalanguages
    'into variable "colDataLang"
    Set colDataLang = Application.AvailableDataLanguages
    iCounter = 1
'
    'Get every languagename and the assigned ID
    For Each objDataLang In colDataLang
        With objDataLang
            If 0 = iCounter Mod 3 Or 1 = iCounter Then
                strLangList = strLangList & vbCrLf & .LanguageID & " " & .LanguageName
            Else
                strLangList = strLangList & " / " & .LanguageID & " " & .LanguageName
            End If
        End With
        iCounter = iCounter + 1
    Next objDataLang
    MsgBox strLangList
End Sub
