Sub ShowLanguageFont()
'VBA278
    Dim colLanguageFonts As HMILanguageFonts
    Dim objLanguageFont As HMILanguageFont
    Dim objButton As HMIButton
    Dim iMax As Integer
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    Set colLanguageFonts = objButton.LDFonts
    iMax = colLanguageFonts.Count
    For Each objLanguageFont In colLanguageFonts
        MsgBox "Planned fonts: " & iMax & vbCrLf & "Language-ID: " & objLanguageFont.LanguageID
    Next objLanguageFont
End Sub
