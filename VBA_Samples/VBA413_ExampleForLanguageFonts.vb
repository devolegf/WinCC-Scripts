Sub ExampleForLanguageFonts()
'VBA413
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "Displaytext"
    Set colLangFonts = objButton.LDFonts
    'Set french fontproperties:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
    'Set english fontproperties:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub
