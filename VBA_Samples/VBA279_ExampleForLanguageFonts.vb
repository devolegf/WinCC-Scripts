Sub ExampleForLanguageFonts()
'VBA279
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "DefText"
    Set colLangFonts = objButton.LDFonts
    
    'Adjust fontsettings for french:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
    'Adjust fontsettings for english:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub
