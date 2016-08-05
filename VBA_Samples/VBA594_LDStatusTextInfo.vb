Sub LDStatusTextInfo()
'VBA594
    Dim colMenuItems As HMIMenuItems
    Dim objMenuItem As HMIMenuItem
    Dim colStatusLngTexts As HMILanguageTexts
    Dim objStatusLngText As HMILanguageText
    Dim strResult As String
    Dim iAnswer As Integer
    Set colMenuItems = ActiveDocument.CustomMenus("DeleteObjects").MenuItems
    For Each objMenuItem In colMenuItems
        strResult = "Statustexts of menuitem """ & objMenuItem.Label & """"
        Set colStatusLngTexts = objMenuItem.LDStatusTexts
        For Each objStatusLngText In colStatusLngTexts
            strResult = strResult & vbCrLf & objStatusLngText.DisplayName
        Next objStatusLngText
        iAnswer = MsgBox(strResult, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objMenuItem
End Sub
