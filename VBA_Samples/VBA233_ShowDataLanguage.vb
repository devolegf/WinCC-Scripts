Sub ShowDataLanguage()
'VBA233
    Dim colDataLanguages As HMIDataLanguages
    Dim objDataLanguage As HMIDataLanguage
    Dim strLanguages As String
    Dim iCount As Integer
    iCount = 0
    Set colDataLanguages = Application.AvailableDataLanguages
    For Each objDataLanguage In colDataLanguages
        If "" <> strLanguages Then strLanguages = strLanguages & "/"
        strLanguages = strLanguages & objDataLanguage.LanguageName & " "
        'Every 15 items of datalanguages output in a messagebox
        If 0 = iCount Mod 15 And 0 <> iCount Then
            MsgBox strLanguages
            strLanguages = ""
        End If
        iCount = iCount + 1
    Next objDataLanguage
    MsgBox strLanguages
End Sub
