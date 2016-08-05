Sub ShowLanguageFont()
'VBA589
    Dim colLanguageFonts As HMILanguageFonts
    Dim objLanguageFont As HMILanguageFont
    Dim objButton As HMIButton
    Dim iMax As Integer
    Dim iAnswer As Integer
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    Set colLanguageFonts = objButton.LDFonts
    iMax = colLanguageFonts.Count
    For Each objLanguageFont In colLanguageFonts
        iAnswer = MsgBox("Projected fonts: " & iMax & vbCrLf & "Language-ID: " & objLanguageFont.LanguageID, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLanguageFont
End Sub
