Sub EditStaticText()
'VBA338
    Dim objStaticText As HMIStaticText
    Set objStaticText = ActiveDocument.HMIObjects("Static_Text1")
    objStaticText.BorderColor = RGB(255, 0, 0)
End Sub
