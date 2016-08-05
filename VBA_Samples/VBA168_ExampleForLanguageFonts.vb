Sub ExampleForLanguageFonts()
'VBA168
    Dim objLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    objButton.Text = "Hello"
    Set objLangFonts = objButton.LDFonts
'
    'To make fontsettings for french:
    With objLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
'
    'To make fontsettings for english:
    With objLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
End Sub
