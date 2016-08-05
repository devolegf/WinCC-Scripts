Sub AddLanguagesToButton()
'VBA2
    Dim objLabelText As HMILanguageText
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
'
    'Set defaultlabel:
    objButton.Text = "Default-Text"
'
    'Add english label:
    Set objLabelText = objButton.LDTexts.Add(1033, "English Text")
    'Add german label:
    Set objLabelText = objButton.LDTexts.Add(1031, "German Text")
End Sub
