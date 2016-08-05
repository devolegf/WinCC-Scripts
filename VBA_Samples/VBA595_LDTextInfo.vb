Sub LDTextInfo()
'VBA595
    Dim colLDLngTexts As HMILanguageTexts
    Dim objLDLngText As HMILanguageText
    Dim objButton As HMIButton
    Dim iAnswer As Integer
    Set objButton = ActiveDocument.HMIObjects("myButton")
    Set colLDLngTexts = objButton.LDTexts
    For Each objLDLngText In colLDLngTexts
        iAnswer = MsgBox(objLDLngText.DisplayName, vbOKCancel)
        If vbCancel = iAnswer Then Exit For
    Next objLDLngText
End Sub
