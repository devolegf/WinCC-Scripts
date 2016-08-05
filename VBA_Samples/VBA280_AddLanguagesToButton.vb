Sub AddLanguagesToButton()
'VBA280
    Dim objLabelText As HMILanguageText
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    '
    'Add text in actual datalanguage:
    objButton.Text = "Actual-Language Text"
    '
    'Add english text:
    Set objLabelText = ActiveDocument.HMIObjects("myButton").LDTexts.Add(1033, "English Text")
End Sub
