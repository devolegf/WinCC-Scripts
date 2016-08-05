Sub ExampleForLanguageFonts()
'VBA492
    Dim colLangFonts As HMILanguageFonts
    Dim objButton As HMIButton
    Dim iStartLangID As Integer
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    iStartLangID = Application.CurrentDataLanguage
    With objButton
        .Text = "Command"
        .Width = 100
    End With
    Set colLangFonts = objButton.LDFonts
'
    'To do typesettings for french:
    With colLangFonts.ItemByLCID(1036)
        .Family = "Courier New"
        .Bold = True
        .Italic = False
        .Underlined = True
        .Size = 12
    End With
'
    'To do typesettings for english:
    With colLangFonts.ItemByLCID(1033)
        .Family = "Times New Roman"
        .Bold = False
        .Italic = True
        .Underlined = False
        .Size = 14
    End With
    With objButton
        Application.CurrentDataLanguage = 1036
        .Text = "Command"
        MsgBox "Datalanguage is changed in french"
        Application.CurrentDataLanguage = 1033
        .Text = "Command"
        MsgBox "Datalanguage is changed in english"
        Application.CurrentDataLanguage = iStartLangID
        MsgBox "Datalanguage is changed back to startlanguage."
    End With
End Sub
