Sub ChangeCurrentDataLanguage()
'VBA1
    Application.CurrentDataLanguage = 1033
    MsgBox "The Data language has been changed to english"
    Application.CurrentDataLanguage = 1031
    MsgBox "The Data language has been changed to german"
End Sub